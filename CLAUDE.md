# CLAUDE.md – Visio-Shapes AddIn

## Projekt-Überblick

VSTO-AddIn für Microsoft Visio (C#). Hostet eine WebView2-Instanz mit der Flask-SPA und stellt Drag-Drop-Funktionalität sowie Upload-Dialoge bereit.

Gegenstück: `../server/` – Flask-Web-App. Ihre API-Dokumentation steht in `../server/CLAUDE.md`.

---

## Tech Stack

| Bereich | Technologie |
|---|---|
| Sprache | C# (.NET Framework, VSTO) |
| UI | WinForms + WebView2 (Microsoft.Web.WebView2) |
| HTTP | System.Net.Http (HttpClient) |
| Serialisierung | Newtonsoft.Json |
| Build | Visual Studio (kein dotnet CLI – VSTO-Projekt) |
| Einstellungen | `Properties.Settings.Default` (user-scoped, als JSON) |
| NuGet-Pakete | lokal in `packages/` |

---

## Verzeichnisstruktur

```
addin/
├── VisioAddin.sln
└── VisioAddin/
    ├── ThisAddIn.cs              # Entry Point, hält PanelManager + SettingsHandler
    ├── PanelManager.cs           # Dictionary<WindowId, PanelFrame> – ein Panel pro Visio-Fenster
    ├── PanelFrame.cs             # Win32-Interop: dockt WinForm als Anchor-Window in Visio
    ├── TheForm.cs                # WinForms-Form: hostet WebView2, Login, Search, WebViewDragDrop
    ├── AddinRibbonComponent.cs   # Ribbon-Button "Toggle Panel"
    ├── Handlers/
    │   └── SettingsHandler.cs    # Liest/schreibt ServerSettings aus Properties.Settings.Default
    ├── Models/
    │   ├── Server.cs             # Name, Url, Token
    │   ├── ServerSettings.cs     # List<Server>, CurrentServer
    │   └── Requests.cs           # OnlineShape, OnlineStencil (Request-Payload-Modelle)
    └── Ui/
        ├── FrmContributeShape.cs    # Uploadformular für einzelne Shape
        ├── FrmContributeStencil.cs  # Uploadformular für ganze Schablone
        └── frmSettings.cs           # Server-Liste verwalten (Name, URL, Token)
```

---

## Schlüsselklassen & Abläufe

### `ThisAddIn`
- Hält `_panelManager` und `ServerHandler` (SettingsHandler)
- `TogglePanel()` → öffnet/schließt Panel für `Application.ActiveWindow`
- Feuert `OnContribute`-Event (nach erfolgreichem Shape-Upload → WebView sucht nach dem Shape-Namen)

### `PanelFrame` (Win32-Interop)
- Erstellt Visio Anchor-Window (`visWSDockedRight | visWSAnchorMerged`)
- `SetParent()` / `SetWindowLong()` zum Einbetten der WinForm
- `AddonWindowMergeId` GUID steuert, in welchen Dock-Bereich gemergert wird

### `TheForm` – der Kern
- Hostet `WebView2`
- **Initialisierung**: `CoreWebView2Environment` → `EnsureCoreWebView2Async` → Host-Object registrieren → `Login()`
- **Login**: POST `/token_login` mit `token=<token>` (form-encoded) als WebView-Navigation → Server setzt Session-Cookie
- **Search**: POST `/` mit `search=<term>` als WebView-Navigation
- **WebViewDragDrop**: `[ComVisible(true)]` Inner Class, registriert als `window.chrome.webview.hostObjects.WebViewDragDrop`

### `WebViewDragDrop.DragDropShape(string shapeData)`
```
JS ruft auf → shapeData = JSON { "Format-Name": "base64-Data", ... }
→ C# deserialisiert → baut DataObject → DoDragDrop(dataObject, DragDropEffects.All)
```

### `FrmContributeShape` – Shape hochladen
1. Liest selektiertes Shape / dessen Master
2. Falls kein Master: erstellt temporäres Stencil-Dokument, fügt Shape ein
3. Serialisiert DataObject → Dictionary `{ format: base64 }` (nur "Visio 15.0 Masters" + "Object Descriptor" gefüllt)
4. Exportiert PNG via `Master.Export()`
5. POST `/add_shape` – Multipart: `json` (OnlineShape-JSON) + `image` (PNG)
6. Nach Erfolg: `RaiseEventOnContribute(name)` → WebView sucht nach dem Shape

### `FrmContributeStencil` – Schablone hochladen
1. Listet offene Stencil-Dokumente (`visTypeStencil`)
2. Iteriert alle Masters → DataObject + PNG je Master
3. POST `/add_stencil` – Multipart: `stencil` (`.vss`-Datei), `json` (OnlineStencil-JSON), `images` (`1.png`, `2.png`, ...)

---

## API-Calls an den Server

| Zweck | Methode | Endpoint | Auth |
|---|---|---|---|
| Login (WebView) | POST | `/token_login` | Body: `token=<token>` |
| Shape hochladen | POST | `/add_shape` | Bearer Token |
| Stencil hochladen | POST | `/add_stencil` | Bearer Token |
| Shape abrufen (JS) | GET | `/get_shape/{id}` | Session-Cookie (WebView) |
| Stencil-Download (JS) | GET | `/download_stencil/{id}` | Session-Cookie (WebView) |
| Suche (WebView) | POST | `/` | Session-Cookie (WebView) |

**Token**: aus `SettingsHandler.CurrentServerToken` – im Server als `User.token` gespeichert.

---

## Einstellungen (Persistenz)

`Properties.Settings.Default.ServerSettings` – user-scoped, als JSON-String gespeichert:
```json
{
  "Servers": [
    { "Name": "www.visio-shapes.com", "Url": "https://www.visio-shapes.com", "Token": "..." },
    { "Name": "localhost", "Url": "http://127.0.0.1:5000", "Token": "..." }
  ],
  "CurrentServer": "www.visio-shapes.com"
}
```

Default-Server beim ersten Start: `www.visio-shapes.com` + `localhost`.

---

## Schnittstelle AddIn ↔ Server (Payload-Formate)

### OnlineShape (JSON-Feld in Multipart)
```json
{
  "Name": "Shape-Name",
  "Prompt": "Tooltip-Text",
  "Keywords": "keyword1 keyword2",
  "DataObject": "{\"Visio 15.0 Masters\": \"base64...\", \"Object Descriptor\": \"base64...\", \"andere Formate\": null}"
}
```

### OnlineStencil (JSON-Feld in Multipart)
```json
{
  "FileName": "MyStencil.vssx",
  "Title": "...", "Subject": "...", "Author": "...",
  "Manager": "...", "Company": "...", "Language": "1033",
  "Categories": "...", "Tags": "...", "Comments": "...",
  "Shapes": [ /* Liste von OnlineShape */ ]
}
```

---

## Offene Aufgaben / Bekannte Lücken

- **Teams-Integration**: Server-Panel zeigt Team-Dropdown bereits an (serverseitig fertig). AddIn muss beim Upload noch `getSelectedTeamId()` aus dem WebView abfragen und als Feld mitsenden.
- **Keine Fehlerbehandlung bei WebView-Host-Object**: JS-Aufrufe auf `WebViewDragDrop` schlagen im normalen Browser stumm fehl. Server-seitig fehlt `try/catch` im JS (steht in `server/CLAUDE.md`).
- **Suche nach Upload**: `RaiseEventOnContribute` löst eine Suche nach dem Shape-Namen aus – das funktioniert nur wenn das Panel offen ist.
