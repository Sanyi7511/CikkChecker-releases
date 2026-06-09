use clap::Parser;
use reqwest::blocking::{Client, ClientBuilder};
use reqwest::header::{HeaderMap, HeaderValue, REFERER, USER_AGENT};
use scraper::{Html, Selector};
use serde_json;
use std::collections::HashSet;
use std::fs;
use std::io::{self, BufRead, Write};
use std::path::Path;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Arc;
use std::thread;
use std::time::{Duration, Instant};
use url::Url;

#[derive(Parser, Debug)]
#[command(name = "checker_core")]
struct Args {
    #[arg(long)] base_url: String,
    #[arg(long, default_value = "")] user: String,
    #[arg(long, default_value = "")] password: String,
    #[arg(long)] requires_login: bool,
    #[arg(long)] codes_file: String,
    #[arg(long)] csv_output: String,
    #[arg(long, default_value_t = 0.8)] sleep_seconds: f64,
    #[arg(long, default_value_t = 100)] save_every: usize,
}

// ── Emit JSON lines to stdout (Python reads these) ───────────────────────────
fn emit_json(obj: serde_json::Value) {
    println!("{}", obj);
    let _ = io::stdout().flush();
}
fn log(msg: &str)     { emit_json(serde_json::json!({"kind":"log","msg":msg})); }
fn status(msg: &str)  { emit_json(serde_json::json!({"kind":"status","msg":msg})); }
fn result(code: &str, avail: &str) { emit_json(serde_json::json!({"kind":"result","cikkszam":code,"elerhetoseg":avail})); }
fn progress(cur: usize, tot: usize) { emit_json(serde_json::json!({"kind":"progress","current":cur,"total":tot})); }
fn unknown(code: &str) { emit_json(serde_json::json!({"kind":"unknown","cikkszam":code})); }
fn done_msg()         { emit_json(serde_json::json!({"kind":"done"})); }
fn error_msg(msg: &str) { emit_json(serde_json::json!({"kind":"error","msg":msg})); }

// ── stdin command (Python writes "stop" or "skip") ───────────────────────────
fn read_stdin_cmd() -> String {
    let mut line = String::new();
    io::stdin().lock().read_line(&mut line).ok();
    line.trim().to_lowercase()
}

// ── URL helpers ───────────────────────────────────────────────────────────────
fn ensure_scheme(url: &str) -> String {
    if url.starts_with("http://") || url.starts_with("https://") { url.to_string() }
    else { format!("https://{}", url) }
}

fn trim_to_root(url: &str) -> String {
    if let Ok(p) = Url::parse(url) {
        format!("{}://{}", p.scheme(), p.host_str().unwrap_or(""))
    } else { url.to_string() }
}

fn normalize(text: &str) -> String {
    text.split_whitespace().collect::<Vec<_>>().join(" ").to_lowercase()
}

fn looks_valid(html: &str) -> bool {
    if html.trim().is_empty() { return false; }
    let doc = Html::parse_document(html);
    let sel = Selector::parse("body").unwrap();
    let text: String = doc.select(&sel).flat_map(|e| e.text()).collect::<Vec<_>>().join(" ");
    let norm = normalize(&text);
    if norm.len() < 30 { return false; }
    !["404 not found","403 forbidden","500 internal server error","not found","access denied"]
        .iter().any(|m| norm.contains(m))
}

// ── HTTP client ───────────────────────────────────────────────────────────────
fn build_client(base_url: &str) -> Client {
    let mut headers = HeaderMap::new();
    headers.insert(USER_AGENT, HeaderValue::from_static("Mozilla/5.0"));
    let referer = format!("{}/login", base_url);
    if let Ok(v) = HeaderValue::from_str(&referer) { headers.insert(REFERER, v); }
    ClientBuilder::new()
        .default_headers(headers)
        .cookie_store(true)
        .timeout(Duration::from_secs(30))
        .gzip(true)
        .redirect(reqwest::redirect::Policy::limited(10))
        .build()
        .expect("HTTP client build failed")
}

// ── Login ─────────────────────────────────────────────────────────────────────
fn login(client: &Client, base_url: &str, user: &str, password: &str) -> bool {
    let login_url = format!("{}/login", base_url);
    log(&format!("Login: {}", login_url));

    let r1 = match client.get(&login_url).send() {
        Ok(r) => r,
        Err(e) => { error_msg(&format!("Login oldal hiba: {}", e)); return false; }
    };
    if r1.status() != 200 { error_msg("Login oldal nem elérhető."); return false; }

    let page_url = r1.url().to_string();
    let html = r1.text().unwrap_or_default();
    let doc = Html::parse_document(&html);

    let form_sel = Selector::parse("form").unwrap();
    let input_sel = Selector::parse("input").unwrap();

    let mut action_url = page_url.clone();
    let mut form_data: Vec<(String, String)> = Vec::new();

    'forms: for form in doc.select(&form_sel) {
        let inputs: Vec<(String, String)> = form.select(&input_sel)
            .filter_map(|i| {
                let name = i.value().attr("name")?;
                Some((name.to_string(), i.value().attr("value").unwrap_or("").to_string()))
            })
            .collect();
        let names: Vec<&str> = inputs.iter().map(|(n,_)| n.as_str()).collect();
        if names.contains(&"user-name") || names.contains(&"password") ||
           form.value().attr("method").unwrap_or("").to_lowercase() == "post" {
            if let Some(action) = form.value().attr("action") {
                if !action.is_empty() {
                    action_url = if action.starts_with("http") { action.to_string() }
                    else { format!("{}/{}", base_url.trim_end_matches('/'), action.trim_start_matches('/')) };
                }
            }
            form_data = inputs;
            break 'forms;
        }
    }

    let mut found_user = false;
    let mut found_pass = false;
    for (k,v) in form_data.iter_mut() {
        if k == "user-name" { *v = user.to_string(); found_user = true; }
        if k == "password"  { *v = password.to_string(); found_pass = true; }
    }
    if !found_user { form_data.push(("user-name".into(), user.into())); }
    if !found_pass { form_data.push(("password".into(), password.into())); }
    form_data.push(("tz-placeholder".into(), "Europe/Budapest".into()));

    if let Err(e) = client.post(&action_url).form(&form_data).send() {
        error_msg(&format!("Login POST hiba: {}", e)); return false;
    }

    let main_url = format!("{}/main", base_url);
    match client.get(&main_url).send() {
        Ok(r) if !r.url().as_str().contains("/login") => { log("Sikeres bejelentkezés."); true }
        _ => { error_msg("Sikertelen bejelentkezés."); false }
    }
}

// ── Stock decision ────────────────────────────────────────────────────────────
fn decide_stock(html: &str, text: &str) -> &'static str {
    let norm = normalize(text);
    let hl = html.to_lowercase();

    let has_price    = norm.contains("ft");
    let has_qty      = text.contains('(') && text.contains(')');
    let has_delivery = ["ma","holnap","holnapután","hétfő","kedd","szerda","csütörtök","péntek","szombat","vasárnap"]
                        .iter().any(|d| norm.contains(d));
    let has_cart     = hl.contains("cart-plus") || norm.contains("kosárba helyezés");
    let disabled     = hl.contains("disabled=\"disabled\"") || hl.contains("pointer-events:none");
    let external     = ["külső raktár","külső készlet","rendelésre","beszerzés alatt"]
                        .iter().any(|k| norm.contains(k));

    if (has_price && has_delivery) || (has_qty && has_price) || (has_cart && !disabled) { return "Van"; }
    if external && !disabled { return "Külső raktár"; }
    if norm.contains("nincs készleten") || norm.contains("nincs raktáron") || disabled {
        if !(has_qty || has_price || has_delivery) { return "Nincs"; }
    }
    "Ismeretlen"
}

fn check_stock(client: &Client, base_url: &str, cikkszam: &str) -> &'static str {
    let url = format!("{}/product-search/{}/0?1", base_url, cikkszam);
    let resp = match client.get(&url).send() {
        Ok(r) => r,
        Err(e) => { log(&format!("Hálózati hiba ({}): {}", cikkszam, e)); return "Ismeretlen"; }
    };
    if resp.status() != 200 { log(&format!("HTTP {} – {}", resp.status(), cikkszam)); return "Ismeretlen"; }

    let html = resp.text().unwrap_or_default();
    if !looks_valid(&html) { log(&format!("Érvénytelen oldal: {}", cikkszam)); return "Ismeretlen"; }

    let doc = Html::parse_document(&html);
    let whole_text: String = doc.root_element().text().collect::<Vec<_>>().join(" ");
    let norm_whole = normalize(&whole_text);

    let no_result = ["nincs találat","nem található","nincs ilyen termék"];
    if no_result.iter().any(|k| norm_whole.contains(k)) && !norm_whole.contains(&cikkszam.to_lowercase()) {
        return "Nincs";
    }

    if norm_whole.contains(&cikkszam.to_lowercase()) {
        decide_stock(&html, &whole_text)
    } else {
        log(&format!("Nem találtam a cikkszámot az oldalon: {}", cikkszam));
        "Ismeretlen"
    }
}

// ── CSV helpers ───────────────────────────────────────────────────────────────
fn load_csv(path: &str) -> (Vec<(String, String)>, HashSet<String>) {
    let mut rows = Vec::new();
    let mut done = HashSet::new();
    if !Path::new(path).exists() { return (rows, done); }
    let mut rdr = match csv::ReaderBuilder::new().has_headers(true).from_path(path) {
        Ok(r) => r, Err(_) => return (rows, done),
    };
    for rec in rdr.records().flatten() {
        let c = rec.get(0).unwrap_or("").trim().to_string();
        let a = rec.get(1).unwrap_or("").trim().to_string();
        if !c.is_empty() { done.insert(c.clone()); rows.push((c, a)); }
    }
    (rows, done)
}

fn write_csv(path: &str, rows: &[(String, String)]) {
    if let Ok(mut w) = csv::Writer::from_path(path) {
        let _ = w.write_record(&["Cikkszám","Elérhetőség"]);
        for (c,a) in rows { let _ = w.write_record(&[c.as_str(), a.as_str()]); }
    }
}

fn append_csv(path: &str, cikkszam: &str, avail: &str) {
    let exists = Path::new(path).exists();
    if let Ok(file) = fs::OpenOptions::new().create(true).append(true).open(path) {
        let mut w = csv::WriterBuilder::new().has_headers(false).from_writer(file);
        if !exists { let _ = w.write_record(&["Cikkszám","Elérhetőség"]); }
        let _ = w.write_record(&[cikkszam, avail]);
    }
}

// ── Main ──────────────────────────────────────────────────────────────────────
fn main() {
    let args = Args::parse();

    let raw = fs::read_to_string(&args.codes_file).unwrap_or_default();
    let mut seen = HashSet::new();
    let unique_codes: Vec<String> = raw.lines()
        .map(|l| l.trim().to_string())
        .filter(|l| !l.is_empty() && seen.insert(l.clone()))
        .collect();

    let total = unique_codes.len();
    if total == 0 { error_msg("Nincs megadva egyetlen cikkszám sem."); return; }
    log(&format!("Egyedi cikkszámok: {}", total));

    let base_url = ensure_scheme(&args.base_url);
    let client = build_client(&base_url);

    status("Oldal ellenőrzése...");
    match client.get(&base_url).send() {
        Ok(r) if r.status() == 200 => {
            if !looks_valid(&r.text().unwrap_or_default()) {
                let root = trim_to_root(&base_url);
                log(&format!("Root URL próba: {}", root));
                match client.get(&root).send() {
                    Ok(r2) if r2.status() == 200 => {}
                    _ => { error_msg("Az oldal nem érhető el."); return; }
                }
            }
        }
        Ok(r) => { error_msg(&format!("HTTP {}", r.status())); return; }
        Err(e) => { error_msg(&format!("Kapcsolódási hiba: {}", e)); return; }
    }

    if args.requires_login {
        status("Bejelentkezés...");
        if !login(&client, &base_url, &args.user, &args.password) {
            status("Login hiba"); return;
        }
    }

    let (mut results, done_set) = load_csv(&args.csv_output);
    let valid_done: HashSet<String> = done_set.intersection(&seen).cloned().collect();
    log(&format!("Már kész: {} / {}", valid_done.len(), total));

    if valid_done.len() == total {
        log("Minden cikkszám feldolgozva.");
        status("Kész"); progress(total, total); done_msg(); return;
    }

    // Stop listener thread
    let stop = Arc::new(AtomicBool::new(false));
    let stop_t = Arc::clone(&stop);
    // NOTE: stop via stdin "stop" line is handled in the unknown branch below;
    // background thread sets the flag for normal iteration.
    thread::spawn(move || {
        for line in io::stdin().lock().lines().flatten() {
            if line.trim() == "stop" { stop_t.store(true, Ordering::SeqCst); break; }
        }
    });

    let sleep_dur = Duration::from_millis((args.sleep_seconds * 1000.0) as u64);

    for (i, code) in unique_codes.iter().enumerate() {
        if stop.load(Ordering::SeqCst) {
            status("Leállítva"); log("Leállítva."); break;
        }
        if valid_done.contains(code) { progress(i+1, total); continue; }

        status(&format!("Feldolgozás: {}", code));
        log(&format!("[{}/{}] {}", i+1, total, code));

        let t0 = Instant::now();
        let avail = check_stock(&client, &base_url, code);
        log(&format!("  → {} ({} ms)", avail, t0.elapsed().as_millis()));

        if avail == "Ismeretlen" {
            unknown(code);
            // Python writes "skip" or "stop" back on stdin
            let cmd = read_stdin_cmd();
            if cmd == "stop" { status("Leállítva"); break; }
            // skip: continue
            progress(i+1, total);
            thread::sleep(sleep_dur);
            continue;
        }

        append_csv(&args.csv_output, code, avail);
        results.push((code.clone(), avail.to_string()));
        result(code, avail);
        progress(i+1, total);

        if (i+1) % args.save_every == 0 { write_csv(&args.csv_output, &results); }
        thread::sleep(sleep_dur);
    }

    write_csv(&args.csv_output, &results);
    status("Kész"); log("Minden kész."); done_msg();
}
