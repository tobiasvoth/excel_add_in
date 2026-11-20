# Excel-Add-in „Schreibe Zwei“

Ein Button im Menüband (Start-Tab → Gruppe „Schreibe Zwei“ → „2 eintragen“), der eine **2** in die aktuell ausgewählte Zelle schreibt.

## Dateien

| Datei | Zweck |
|---|---|
| `manifest.xml` | Beschreibt das Add-in: Name, Button, Berechtigungen, URLs |
| `public/commands.html` + `commands.js` | Die Funktion `writeTwo`, die der Button ausführt |
| `public/taskpane.html` | Pflicht-Platzhalterseite für das Manifest |
| `public/assets/icon-*.png` | Icons für den Button |
| `server.js` | Lokaler HTTPS-Server (nur zum Entwickeln) |
| `package.json` | Abhängigkeiten und Start-Skripte |

## Lokal testen (Windows oder Mac, Excel Desktop)

Voraussetzung: [Node.js](https://nodejs.org) (LTS).

```bash
npm install
npx office-addin-dev-certs install    # Entwicklungszertifikat einmalig vertrauen (Rückfrage bestätigen)
npm start                             # startet Server und öffnet Excel mit dem Add-in
```

In Excel: Zelle anklicken → Start-Tab → **2 eintragen**.
Beenden mit `npm run stop`.

## Excel im Browser (ohne npm start)

1. `npm run dev-server` in einem Terminal laufen lassen (Zertifikat vorher wie oben installieren).
2. Excel im Browser öffnen → *Einfügen → Add-Ins → Meine Add-Ins verwalten → Mein Add-In hochladen* → `manifest.xml` wählen.

## Für mehrere Nutzer bereitstellen

1. Den Ordner `public` auf einen HTTPS-Webspace hochladen (z. B. `https://addin.meinefirma.de/`).
2. In `manifest.xml` alle `https://localhost:3000` durch diese URL ersetzen.
3. Optional neue `<Id>` (GUID) vergeben, wenn du eine eigene Variante veröffentlichst.
4. Manifest im Microsoft 365 Admin Center hochladen (*Einstellungen → Integrierte Apps → Benutzerdefinierte Apps hochladen*) und Nutzern/Gruppen zuweisen.

## Manifest prüfen

```bash
npm run validate
```
