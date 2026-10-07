# CLAUDE.md – Dienstplan Assistenz

## Zweck
Web-Ansicht (PWA) des Dienstplans der Assistenzärzt:innen. Die Pläne werden weiterhin in Excel erstellt; veröffentlicht wird nur ein verschlüsseltes Bundle (`plaene/plan.enc`), das im Browser mit dem Team-Passwort entschlüsselt wird.

## Stack
- Reines statisches HTML/CSS/JavaScript, kein Build-Schritt, kein Paketmanager, keine Tests.
- Excel-Parsing im Browser über SheetJS (`xlsx@0.18.5` von cdn.jsdelivr.net), Schrift von Google Fonts.
- Verschlüsselung per Web Crypto: AES-256-GCM, Schlüssel per PBKDF2-SHA256 (310 000 Iterationen) aus dem Team-Passwort.
- Service Worker (`sw.js`): Network-first, offline letzte Version aus dem Cache.

## Struktur
- `index.html` – Seite; bindet `style.css?v=N` und `app.js?v=N` ein (Cache-Busting).
- `app.js` – gesamte Logik: Konfiguration (Kürzel, Spalten, Status), Excel-Parser, Ver-/Entschlüsselung, Ansichten (Meine Dienste, Monatsplan, Tag, Jahr), Export (.ics, .xlsx, PDF), Prüfungen für die Planerin.
- `style.css` – Styles inkl. Druck-Layout.
- `sw.js`, `manifest.webmanifest`, `*.png` – PWA/Icons.
- `plaene/plan.enc` – verschlüsseltes JSON-Bundle (Excel-Dateien + optionale Korrektur-JSON).
- `.nojekyll` – GitHub Pages ohne Jekyll.

## Befehle
- Bauen: entfällt (statische Dateien).
- Verschlüsseln/Veröffentlichen: nur über die Webseite selbst – entsperren → „Für die Dienstplanerin“ → Excel hineinziehen → Vorschau/Prüfung → „Verschlüsselt speichern“ lädt `plan.enc` herunter → in `plaene/` ersetzen und committen (siehe README.md).
- Lokal ansehen: über einen beliebigen statischen HTTP-Server im Repo-Ordner (z. B. `python -m http.server 8000`), nicht per `file://` (der `fetch` von `plan.enc` scheitert sonst). Der Service Worker registriert sich nur unter https.
- Nach Änderungen an `app.js`/`style.css` die `?v=`-Nummer in `index.html` erhöhen.

## Deployment
- GitHub Pages (legacy, Quelle: Branch `main`, Root) → https://dendak.github.io/dienstplan/
- Keine GitHub Actions. **Jeder Push/Commit auf `main` ist sofort live** (ca. 1 Minute), auch Uploads über die GitHub-Weboberfläche.

## Vorsicht
- Öffentliches Repo. Klartext-Dienstplandaten (Excel `*.xlsx/*.xlsm/*.xls`, entschlüsseltes JSON, Exporte) **nie committen**.
- Nur `plaene/plan.enc` (verschlüsselt) gehört ins Repo. Nichts entschlüsseln und das Ergebnis speichern.
- Team-Passwort nie ins Repo, in Code, Commit-Messages, Notizen oder Issues schreiben.
- `.gitignore` (Excel-Dateien, `PASSWORT*`) nicht aufweichen.
- Format von `plan.enc` und Krypto-Parameter nicht ohne Not ändern – sonst kann die bestehende Datei nicht mehr gelesen werden.
- `app.js` enthält Kürzel→Namen-Zuordnungen in Klartext (öffentlich sichtbar); keine weiteren Personendaten in den Code oder Commit-Messages aufnehmen.
- Push auf `main` = Live-Deployment für das ganze Team.

## Arbeit über mehrere Geräte
- GitHub (`Dendak/dienstplan`) ist die Quelle der Wahrheit.
- Session-Start: `git pull`, dann diese Datei und [docs/POZNAMKY.md](docs/POZNAMKY.md) lesen.
- Session-Ende: `docs/POZNAMKY.md` aktualisieren, committen, pushen.
- Größere oder riskante Änderungen über Branch + Pull Request, weil `main` direkt live deployt.
- Repo lokal unter `C:\Users\holub\code\dienstplan`, nie in OneDrive.
- Die Excel-Quelldateien liegen nicht im Repo (per `.gitignore` ausgeschlossen) und müssen auf einem neuen Gerät separat besorgt werden. Das Team-Passwort ebenfalls nur außerhalb des Repos.
