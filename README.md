# Dienstplan Assistenz

Web-Ansicht des Dienstplans der Assistenzärzt:innen. Die Pläne werden weiterhin in Excel erstellt.

- `plaene/plan.enc` – alle veröffentlichten Pläne, **verschlüsselt** (AES-256-GCM, Schlüssel per PBKDF2 aus dem Team-Passwort). Ohne Passwort ist der Inhalt nicht lesbar.
- **Nie Excel-Dateien direkt in dieses Repository hochladen.**

## Neuen Plan veröffentlichen

1. Webseite öffnen, mit dem Team-Passwort entsperren.
2. Unten „Für die Dienstplanerin“ → Excel-Datei hineinziehen → Vorschau & Prüfung.
3. „Verschlüsselt speichern“ → `plan.enc` wird heruntergeladen.
4. Hier im Ordner `plaene` hochladen (bestehende Datei ersetzen) → „Commit changes“.

Nach ca. 1 Minute ist der neue Plan online.
