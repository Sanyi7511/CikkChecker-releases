# CikkChecker

Automatikus cikkszám-elérhetőség ellenőrző a Szakál Metal webshophoz.

## Letöltés / Telepítés

👉 **[Legújabb verzió letöltése](https://github.com/Sanyi7511/CikkChecker-releases/releases/latest)**

Töltsd le a `CikkCheckerSetup.exe` fájlt, futtasd, és kész.

## Rendszerkövetelmény

- Windows 10 / 11 (64-bit)

## Fejlesztés

### Struktúra
```
app.py               ← Python UI (CustomTkinter)
checker_core_src/    ← Rust HTTP backend
installer.iss        ← Inno Setup telepítő script
.github/workflows/   ← Automatikus build
```

### Új verzió kiadása
1. Módosítsd az `APP_VERSION`-t az `app.py`-ban
2. Commitolj és pusholj
3. Hozz létre egy új taget: `git tag v2.1.0 && git push origin v2.1.0`
4. A GitHub Actions automatikusan elkészíti és feltölti a `CikkCheckerSetup.exe`-t
