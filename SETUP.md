# Zenaris Outlook Add-in: Setup Anleitung

## Was ist das?

Ein Outlook Button der per Knopfdruck einen AI Draft erstellt.
Du wählst den Modus (Antwort / mit Termin / nur Termin), optional gibst du extra Kontext,
und der Draft landet in deinen Entwürfen.

## Dateien im Paket

```
zenaris-outlook-addin/
├── manifest.xml          ← Add-in Definition für Outlook
├── taskpane.html         ← Das Panel mit den Buttons (wird in Outlook angezeigt)
├── icon-16.png           ← Icons in verschiedenen Größen
├── icon-32.png
├── icon-64.png
├── icon-80.png
├── icon-128.png
└── SETUP.md              ← Diese Anleitung
```

## Schritt 1: GitHub Repository erstellen

1. Geh zu https://github.com/new
2. Repository Name: `zenaris-outlook-addin`
3. Auf **Public** setzen (GitHub Pages braucht das bei Free Accounts)
4. "Create repository" klicken
5. Alle Dateien aus diesem Ordner hochladen (Upload files)

## Schritt 2: GitHub Pages aktivieren

1. Im Repository: Settings → Pages
2. Source: "Deploy from a branch"
3. Branch: `main`, Ordner: `/ (root)`
4. Save klicken
5. Warte 1 bis 2 Minuten, dann ist deine Seite live unter:
   `https://DEIN_USERNAME.github.io/zenaris-outlook-addin/`

## Schritt 3: URLs in den Dateien anpassen

In **manifest.xml** alle Stellen mit `DEIN_GITHUB_USERNAME` ersetzen:

Suche nach: `DEIN_GITHUB_USERNAME.github.io`
Ersetze mit: `dein-echter-username.github.io`

(Kommt 6 mal vor in der Datei)

In **taskpane.html** die Webhook URL eintragen:

Suche nach: `DEINE_N8N_WEBHOOK_URL_HIER`
Ersetze mit: Die URL vom Webhook Node aus deinem n8n Workflow

## Schritt 4: n8n Workflow vorbereiten

1. Dupliziere deinen bestehenden E-Mail Draft Workflow
2. Lösche den "Microsoft Outlook Trigger" Node
3. Füge stattdessen einen **Webhook** Node ein:
   - HTTP Method: POST
   - Path: z.B. `zenaris-draft`
   - Response Mode: "Last Node"
4. Die Webhook URL die n8n dir anzeigt (z.B. `https://dein-n8n.com/webhook/zenaris-draft`)
   trägst du in die taskpane.html ein (siehe Schritt 3)

### Workflow Anpassungen:

Der Webhook liefert dir folgende Daten im Body:
- `mode`: "reply", "reply_calendar", oder "calendar_only"
- `itemId`: Die Outlook Mail ID
- `from`: Absender E-Mail
- `fromName`: Absender Name
- `subject`: Betreff
- `body`: Mail Text (plain text)
- `extraContext`: Optionaler Kontext vom User

**"Get Full Mail" Node anpassen:**
URL ändern zu:
```
https://graph.microsoft.com/v1.0/me/messages/{{ $json.body.itemId }}
```
(Weil die Mail ID jetzt vom Webhook kommt, nicht vom Trigger)

**Filter Nodes:**
Können entfernt oder deaktiviert werden! Du triggerst ja bewusst per Knopfdruck,
also brauchst du keine Filter mehr.

**Code Node anpassen:**
Am Anfang den `mode` und `extraContext` auslesen:

```javascript
const events = $input.all();
const trigger = $('Get Full Mail').first().json;
const webhookData = $('Webhook').first().json.body;
const mode = webhookData.mode;
const extraContext = webhookData.extraContext || "";

// Bei "calendar_only" keinen inhaltlichen Draft
// Bei "reply" keinen Kalender Check
// Bei "reply_calendar" beides
```

**System Prompt im OpenAI Node anpassen:**
Den `mode` und `extraContext` in den Prompt einbauen:

```
MODUS: {{ $json.mode }}
- "reply" = Normale Antwort, KEINE Terminvorschläge
- "reply_calendar" = Antwort MIT Terminvorschlägen aus dem Kalender
- "calendar_only" = NUR Terminvorschläge, kein inhaltlicher Draft

EXTRA KONTEXT VOM USER: {{ $json.extraContext }}
Falls Extra Kontext vorhanden, berücksichtige diesen bei der Antwort.
z.B. "sag ab" = Absage formulieren, "zusagen" = Zusage formulieren
```

## Schritt 5: Add-in in Outlook installieren

1. Geh im Browser zu: https://aka.ms/olksideload
2. Outlook Web öffnet sich mit dem Add-ins Dialog
3. Wähle "Meine Add-Ins"
4. Unten: "Benutzerdefinierte Add-Ins" → "Benutzerdefiniertes Add-In hinzufügen" → "Aus Datei hinzufügen"
5. Lade deine `manifest.xml` hoch
6. Bestätige die Installation

Das Add-in erscheint dann in deiner Outlook Toolbar als "Draft erstellen" Button.
Es funktioniert automatisch auch in Outlook für Mac und Outlook Desktop.

## Testen

1. Öffne eine beliebige E-Mail in Outlook
2. Klicke auf den "Draft erstellen" Button in der Toolbar
3. Das Zenaris Panel öffnet sich rechts
4. Wähle einen Modus
5. Optional: Extra Kontext eingeben
6. Klicke "Draft generieren"
7. Check deine Entwürfe!

## Troubleshooting

**Button erscheint nicht?**
→ Outlook neu starten, kann bis zu 24h dauern bei Desktop (bei Web sofort)

**"Draft konnte nicht generiert werden"?**
→ Webhook URL in taskpane.html prüfen
→ n8n Workflow prüfen ob er aktiv ist

**Panel zeigt "Mail konnte nicht geladen werden"?**
→ Add-in braucht ReadItem Permission, sollte aber schon gesetzt sein in manifest.xml
