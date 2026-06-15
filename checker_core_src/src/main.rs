use clap::Parser;
use reqwest::blocking::{Client, ClientBuilder};
use reqwest::header::{HeaderMap, HeaderValue, REFERER, USER_AGENT};
use scraper::{Html, Selector};
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
    #[arg(long, default_value = "abc")] sort_order: String,
}

// ── JSON emit ────────────────────────────────────────────────────────────────
fn emit(obj: serde_json::Value) { println!("{}", obj); let _ = io::stdout().flush(); }
fn log(msg: &str)      { emit(serde_json::json!({"kind":"log","msg":msg})); }
fn status(msg: &str)   { emit(serde_json::json!({"kind":"status","msg":msg})); }
fn progress(cur: usize, tot: usize) { emit(serde_json::json!({"kind":"progress","current":cur,"total":tot})); }
fn unknown(code: &str) { emit(serde_json::json!({"kind":"unknown","cikkszam":code})); }
fn done_msg()          { emit(serde_json::json!({"kind":"done"})); }
fn error_msg(msg: &str){ emit(serde_json::json!({"kind":"error","msg":msg})); }
fn login_detected(required: bool) { emit(serde_json::json!({"kind":"login_detected","required":required})); }

fn result_full(code: &str, avail: &str, price: &str, images: usize) {
    emit(serde_json::json!({"kind":"result","cikkszam":code,"elerhetoseg":avail,"ar":price,"images":images}));
}

// ── stdin ────────────────────────────────────────────────────────────────────
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

fn count_product_images(html: &str, cikkszam: &str) -> usize {
    let doc = Html::parse_document(html);
    // Try to find product block first
    if let Ok(sel) = Selector::parse(&format!("[title='{}']", cikkszam)) {
        if let Some(el) = doc.select(&sel).next() {
            // Walk up ~6 levels to find product block
            let block_html = format!("{:?}", el.html());
            // Count img tags in surrounding HTML — approximate by searching in whole page
            // since scraper doesn't easily let us walk upward
        }
    }
    // Count ALL img tags in the page, excluding tiny icons (width/height hints)
    if let Ok(img_sel) = Selector::parse("img") {
        let imgs: Vec<_> = doc.select(&img_sel).collect();
        // Filter out icons: only count imgs that likely have meaningful size
        // We check for src patterns typical of product images
        let product_imgs = imgs.iter().filter(|img| {
            let src = img.value().attr("src").unwrap_or("");
            let alt = img.value().attr("alt").unwrap_or("");
            let class = img.value().attr("class").unwrap_or("");
            // Skip flags, icons, logos
            !src.contains("flag") && !src.contains("icon") && !src.contains("logo")
            && !class.contains("flag") && !class.contains("icon")
            && (src.contains("product") || src.contains("item") || src.contains("upload")
                || src.ends_with(".jpg") || src.ends_with(".png") || src.ends_with(".webp")
                || !alt.is_empty())
        }).count();
        return product_imgs;
    }
    0
}

fn looks_valid(html: &str) -> bool {
    if html.trim().is_empty() { return false; }
    let doc = Html::parse_document(html);
    if let Ok(sel) = Selector::parse("body") {
        let text: String = doc.select(&sel).flat_map(|e| e.text()).collect::<Vec<_>>().join(" ");
        let norm = normalize(&text);
        if norm.len() < 30 { return false; }
        return !["404 not found","403 forbidden","500 internal server error","not found","access denied"]
            .iter().any(|m| norm.contains(m));
    }
    false
}

// ── HTTP client ───────────────────────────────────────────────────────────────
fn build_client(base_url: &str) -> Client {
    let mut headers = HeaderMap::new();
    headers.insert(USER_AGENT, HeaderValue::from_static("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"));
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

// ── Auto-detect login requirement ─────────────────────────────────────────────
fn detect_login_required(client: &Client, base_url: &str) -> bool {
    // Try to access a product search page without login
    let test_url = format!("{}/product-search/test/0?1", base_url);
    match client.get(&test_url).send() {
        Ok(r) => {
            let final_url = r.url().to_string();
            let html = r.text().unwrap_or_default();
            let norm = normalize(&html);
            // Redirected to login page OR page contains login form
            if final_url.contains("/login") { return true; }
            if norm.contains("bejelentkezés") || norm.contains("belépés") ||
               norm.contains("felhasználónév") || norm.contains("jelszó") ||
               norm.contains("login") || norm.contains("sign in") { return true; }
            false
        }
        Err(_) => false,
    }
}

// ── Login ─────────────────────────────────────────────────────────────────────
fn login(client: &Client, base_url: &str, user: &str, password: &str) -> bool {
    let login_url = format!("{}/login", base_url);
    log(&format!("Login: {}", login_url));

    let r1 = match client.get(&login_url).send() {
        Ok(r) => r,
        Err(e) => { error_msg(&format!("Login oldal hiba: {}", e)); return false; }
    };
    if r1.status() != 200 { error_msg("Login oldal nem elerheto."); return false; }

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
        Ok(r) if !r.url().as_str().contains("/login") => { log("Sikeres bejelentkezes."); true }
        _ => { error_msg("Sikertelen bejelentkezes."); false }
    }
}

// ── Price extraction ──────────────────────────────────────────────────────────
fn extract_price(html: &str, text: &str) -> String {
    let doc = Html::parse_document(html);

    // Try common price selectors
    let price_selectors = [
        ".price", ".product-price", ".ar", "[class*='price']",
        "[class*='Price']", "[class*='ar']", "span.netto", "span.brutto",
        ".netto-price", ".brutto-price",
    ];

    for sel_str in &price_selectors {
        if let Ok(sel) = Selector::parse(sel_str) {
            for elem in doc.select(&sel) {
                let t = elem.text().collect::<Vec<_>>().join(" ");
                let cleaned = clean_price(&t);
                if !cleaned.is_empty() { return cleaned; }
            }
        }
    }

    // Fallback: regex-style search in full text for "X Ft" or "X HUF" patterns
    let norm = normalize(text);
    for token in norm.split_whitespace() {
        if let Some(stripped) = token.strip_suffix("ft") {
            let digits: String = stripped.chars().filter(|c| c.is_ascii_digit() || *c == ' ').collect();
            let digits = digits.trim().to_string();
            if !digits.is_empty() && digits.len() >= 3 {
                return format!("{} Ft", digits);
            }
        }
    }

    // Look for price pattern: number followed by "Ft" in text
    let text_lower = text.to_lowercase();
    if let Some(ft_pos) = text_lower.find(" ft") {
        let before = &text[..ft_pos];
        let price_part: String = before.chars().rev()
            .take_while(|c| c.is_ascii_digit() || *c == ' ' || *c == '\u{a0}')
            .collect::<String>()
            .chars().rev()
            .collect();
        let price_clean: String = price_part.chars().filter(|c| c.is_ascii_digit()).collect();
        if price_clean.len() >= 3 {
            return format_price_num(&price_clean);
        }
    }

    String::new()
}

fn clean_price(raw: &str) -> String {
    let norm = raw.trim();
    if norm.is_empty() { return String::new(); }
    // Must contain digits and Ft
    let has_digits = norm.chars().any(|c| c.is_ascii_digit());
    let has_ft = norm.to_lowercase().contains("ft") || norm.contains("HUF");
    if !has_digits { return String::new(); }

    // Extract just number + Ft
    let digits: String = norm.chars().filter(|c| c.is_ascii_digit()).collect();
    if digits.len() < 2 { return String::new(); }

    if has_ft {
        format_price_num(&digits) + " Ft"
    } else {
        format_price_num(&digits)
    }
}

fn format_price_num(digits: &str) -> String {
    // Add thousand separators
    let chars: Vec<char> = digits.chars().collect();
    let mut result = String::new();
    for (i, c) in chars.iter().enumerate() {
        if i > 0 && (chars.len() - i) % 3 == 0 { result.push(' '); }
        result.push(*c);
    }
    result
}

// ── Stock decision ────────────────────────────────────────────────────────────
fn decide_stock(html: &str, text: &str) -> &'static str {
    let norm = normalize(text);
    let hl = html.to_lowercase();

    let has_price    = norm.contains("ft");
    let has_qty      = text.contains('(') && text.contains(')');
    let has_delivery = ["ma ","holnap","holnaputan","hetfo","kedd","szerda","csutortok","pentek","szombat","vasarnap",
                        "hétfő","kedd","szerda","csütörtök","péntek","holnapután"]
                        .iter().any(|d| norm.contains(d));
    let has_cart     = hl.contains("cart-plus") || norm.contains("kosarba helyezes") || norm.contains("kosárba");
    let disabled     = hl.contains("disabled=\"disabled\"") || hl.contains("pointer-events:none");
    let external     = ["kulso raktar","kulső raktár","kulso keszlet","rendelesre","rendel\u{e9}sre","besz\u{e9}rz\u{e9}s"]
                        .iter().any(|k| norm.contains(k));

    if (has_price && has_delivery) || (has_qty && has_price) || (has_cart && !disabled) { return "Van"; }
    if external && !disabled { return "Külső raktár"; }
    if norm.contains("nincs keszleten") || norm.contains("nincs raktaron") ||
       norm.contains("nincs készleten") || norm.contains("nincs raktáron") || disabled {
        if !(has_qty || has_price || has_delivery) { return "Nincs"; }
    }
    "Ismeretlen"
}

fn check_stock(client: &Client, base_url: &str, cikkszam: &str) -> (&'static str, String, usize) {
    let url = format!("{}/product-search/{}/0?1", base_url, cikkszam);
    let resp = match client.get(&url).send() {
        Ok(r) => r,
        Err(e) => { log(&format!("Halozati hiba ({}): {}", cikkszam, e)); return ("Ismeretlen", String::new(), 0); }
    };
    if resp.status() != 200 {
        log(&format!("HTTP {} - {}", resp.status(), cikkszam));
        return ("Ismeretlen", String::new(), 0);
    }

    let html = resp.text().unwrap_or_default();
    if !looks_valid(&html) { return ("Ismeretlen", String::new(), 0); }

    let doc = Html::parse_document(&html);
    let whole_text: String = doc.root_element().text().collect::<Vec<_>>().join(" ");
    let norm_whole = normalize(&whole_text);

    let no_result = ["nincs talalat","nem talalhato","nincs ilyen termek",
                     "nincs találat","nem található","nincs ilyen termék"];
    if no_result.iter().any(|k| norm_whole.contains(k)) && !norm_whole.contains(&cikkszam.to_lowercase()) {
        return ("Nincs", String::new(), 0);
    }

    if norm_whole.contains(&cikkszam.to_lowercase()) {
        let avail = decide_stock(&html, &whole_text);
        let price = if avail != "Nincs" && avail != "Ismeretlen" {
            extract_price(&html, &whole_text)
        } else {
            String::new()
        };
        let images = count_product_images(&html, cikkszam);
        (avail, price, images)
    } else {
        log(&format!("Nem talaltam: {}", cikkszam));
        ("Ismeretlen", String::new(), 0)
    }
}

// ── CSV helpers ───────────────────────────────────────────────────────────────
fn load_csv(path: &str) -> (Vec<(String, String, String, usize)>, HashSet<String>) {
    let mut rows = Vec::new();
    let mut done = HashSet::new();
    if !Path::new(path).exists() { return (rows, done); }
    let mut rdr = match csv::ReaderBuilder::new().has_headers(true).from_path(path) {
        Ok(r) => r, Err(_) => return (rows, done),
    };
    for rec in rdr.records().flatten() {
        let c = rec.get(0).unwrap_or("").trim().to_string();
        let a = rec.get(1).unwrap_or("").trim().to_string();
        let p = rec.get(2).unwrap_or("").trim().to_string();
        let i: usize = rec.get(3).unwrap_or("0").trim().parse().unwrap_or(0);
        if !c.is_empty() { done.insert(c.clone()); rows.push((c, a, p, i)); }
    }
    (rows, done)
}

fn write_csv(path: &str, rows: &[(String, String, String, usize)]) {
    if let Ok(mut w) = csv::Writer::from_path(path) {
        let _ = w.write_record(&["Cikkszam","Elerhetoseg","Ar","Kepek"]);
        for (c,a,p,i) in rows {
            let img_str = i.to_string();
            let _ = w.write_record(&[c.as_str(), a.as_str(), p.as_str(), &img_str]);
        }
    }
}

fn append_csv(path: &str, cikkszam: &str, avail: &str, price: &str, images: usize) {
    let exists = Path::new(path).exists();
    if let Ok(file) = fs::OpenOptions::new().create(true).append(true).open(path) {
        let mut w = csv::WriterBuilder::new().has_headers(false).from_writer(file);
        if !exists { let _ = w.write_record(&["Cikkszam","Elerhetoseg","Ar","Kepek"]); }
        let img_str = images.to_string();
        let _ = w.write_record(&[cikkszam, avail, price, &img_str]);
    }
}

// ── Sort helpers ──────────────────────────────────────────────────────────────
fn sort_codes(codes: &mut Vec<String>, order: &str) {
    match order {
        "abc" => codes.sort_by(|a, b| a.to_lowercase().cmp(&b.to_lowercase())),
        "abc_desc" => codes.sort_by(|a, b| b.to_lowercase().cmp(&a.to_lowercase())),
        "num" => codes.sort_by(|a, b| {
            let na: u64 = a.chars().filter(|c| c.is_ascii_digit()).collect::<String>().parse().unwrap_or(0);
            let nb: u64 = b.chars().filter(|c| c.is_ascii_digit()).collect::<String>().parse().unwrap_or(0);
            na.cmp(&nb)
        }),
        "num_desc" => codes.sort_by(|a, b| {
            let na: u64 = a.chars().filter(|c| c.is_ascii_digit()).collect::<String>().parse().unwrap_or(0);
            let nb: u64 = b.chars().filter(|c| c.is_ascii_digit()).collect::<String>().parse().unwrap_or(0);
            nb.cmp(&na)
        }),
        _ => {} // "none" = no sort, keep original order
    }
}

// ── Main ──────────────────────────────────────────────────────────────────────
fn main() {
    let args = Args::parse();

    let raw = fs::read_to_string(&args.codes_file).unwrap_or_default();
    let mut seen = HashSet::new();
    let mut unique_codes: Vec<String> = raw.lines()
        .map(|l| l.trim().to_string())
        .filter(|l| !l.is_empty() && seen.insert(l.clone()))
        .collect();

    // Sort
    sort_codes(&mut unique_codes, &args.sort_order);
    seen = unique_codes.iter().cloned().collect();

    let total = unique_codes.len();
    if total == 0 { error_msg("Nincs megadva egyetlen cikkszam sem."); return; }
    log(&format!("Egyedi cikkszamok: {} (rendezve: {})", total, args.sort_order));

    let base_url = ensure_scheme(&args.base_url);
    let client = build_client(&base_url);

    // Auto-detect login
    status("Oldal ellenorzese...");
    match client.get(&base_url).send() {
        Ok(r) if r.status() == 200 => {
            if !looks_valid(&r.text().unwrap_or_default()) {
                let root = trim_to_root(&base_url);
                match client.get(&root).send() {
                    Ok(r2) if r2.status() == 200 => {}
                    _ => { error_msg("Az oldal nem erheto el."); return; }
                }
            }
        }
        Ok(r) => { error_msg(&format!("HTTP {}", r.status())); return; }
        Err(e) => { error_msg(&format!("Kapcsolodasi hiba: {}", e)); return; }
    }

    // Detect if login is required
    status("Bejelentkezes detektalasa...");
    let login_required = if args.requires_login {
        true
    } else {
        let detected = detect_login_required(&client, &base_url);
        login_detected(detected);
        detected
    };

    if login_required {
        login_detected(true);
        status("Bejelentkezes...");
        if !login(&client, &base_url, &args.user, &args.password) {
            status("Login hiba"); return;
        }
    } else {
        log("Bejelentkezes nem szukseges.");
    }

    let (mut results, done_set) = load_csv(&args.csv_output);
    let valid_done: HashSet<String> = done_set.intersection(&seen).cloned().collect();
    log(&format!("Mar kesz: {} / {}", valid_done.len(), total));

    if valid_done.len() == total {
        log("Minden cikkszam feldolgozva.");
        status("Kesz"); progress(total, total); done_msg(); return;
    }

    let stop = Arc::new(AtomicBool::new(false));
    let stop_t = Arc::clone(&stop);
    thread::spawn(move || {
        for line in io::stdin().lock().lines().flatten() {
            if line.trim() == "stop" { stop_t.store(true, Ordering::SeqCst); break; }
        }
    });

    let sleep_dur = Duration::from_millis((args.sleep_seconds * 1000.0) as u64);

    for (i, code) in unique_codes.iter().enumerate() {
        if stop.load(Ordering::SeqCst) {
            status("Leallitva"); log("Leallitva."); break;
        }
        if valid_done.contains(code) { progress(i+1, total); continue; }

        status(&format!("Feldolgozas: {}", code));
        log(&format!("[{}/{}] {}", i+1, total, code));

        let t0 = Instant::now();
        let (avail, price) = check_stock(&client, &base_url, code);
        log(&format!("  -> {} | Ar: {} ({} ms)", avail, if price.is_empty() { "-" } else { &price }, t0.elapsed().as_millis()));

        if avail == "Ismeretlen" {
            unknown(code);
            let cmd = read_stdin_cmd();
            if cmd == "stop" { status("Leallitva"); break; }
            progress(i+1, total);
            thread::sleep(sleep_dur);
            continue;
        }

        append_csv(&args.csv_output, code, avail, &price);
        results.push((code.clone(), avail.to_string(), price.clone()));
        result_full(code, avail, &price);
        progress(i+1, total);

        if (i+1) % args.save_every == 0 { write_csv(&args.csv_output, &results); }
        thread::sleep(sleep_dur);
    }

    write_csv(&args.csv_output, &results);
    status("Kesz"); log("Minden kesz."); done_msg();
}
