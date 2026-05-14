# Notion Client
The inspiration for this project came from my own experience using Notion. Though Notion's calendar view is very helpful, it would be even better if it was able to integrate with existing calendar solutions so that my Outlook or Google Calendar contained all the information I need for my day. As a student, my Notion calendar is full of class deadlines months and months in advance and incorpating all those events into my calendar without having to spend the afternoon manually inputting them.

> **Note**: Some of the setup of this project is specific to how my Notion database is setup. This project can be cloned and changed accordingly to account for differences in how notion users have their setups.

## Notion API
Notion's new API is very powerful and has allowed this project to happen. With their excellent development tools, I was able to quickly get this project working and collect the data I needed from my existing Notion databases.

## Workflow
1. Connect to Notion with API Key
2. Get all events from board
3. Organize into dictionary
4. Organize into iCal format
5. Export iCal file

## Setup
### API Token
Go to (Notion API Website)[https://www.notion.so/my-integrations] and register a new integration and receive your integration's API key.

### Attach Database to Integration
1. Go to your board in (Notion)[Notion.so]
2. Click on "Share" in the top right corner
3. Invite your integration to the board

### Git Clone
To clone this repo, please run ...
`git clone https://github.com/radroof22/Notion-iCal`

### Virtual Environment
In order to run the virtual environment, run ...
```bash
source env/bin/activate
```

### Environment Variables
This project makes use of environment variables. In order to edit the configuration for your setup, please add a `.env` file in the root directory. You must set the following environment variables...
`
NOTION_TOKEN=############
DATABASE_ID=##########
`

### Run It!
Just run ...
`python3 main.py`
and see Notion.ics be generated

### Add it!
Import the `Notion.ics` into your favorite calendar app and enjoy!

## Abo-Feed mit GitHub Pages (HTTPS-URL)

So bekommst du eine **URL**, die du in Google Kalender / Outlook etc. unter **„Per URL abonnieren“** eintragen kannst. Die Datei wird **nicht** bei jedem Kalenderabruf live von Notion gezogen, sondern von einem **geplanten GitHub Actions**-Job neu erzeugt und auf **GitHub Pages** veröffentlicht (typisch alle paar Stunden).

### Voraussetzungen

- Repo auf **GitHub** (öffentlich oder mit passendem Plan für private Pages).
- In Notion: Integration wie zuvor, **Datenbank freigegeben**.

### Einrichtung

1. **Repository** mit diesem Code auf GitHub anlegen / pushen.
2. Unter **Settings → Secrets and variables → Actions** zwei **Repository secrets** anlegen:
   - `NOTION_TOKEN` – dein Notion-Integrationstoken  
   - `DATABASE_ID` – dieselbe ID wie in `.env` lokal  
3. **Settings → Pages**: Quelle **„Deploy from a branch“** wählen, Branch **`gh-pages`**, Ordner **`/ (root)`**.  
   *(Der Branch wird beim ersten erfolgreichen Workflow-Lauf automatisch angelegt.)*
4. Unter **Actions** den Workflow **„Publish Notion iCal“** einmal manuell starten (**Run workflow**), warten bis er grün ist.
5. Deine Feed-URL (Projekt-Repository, Beispiel):  
   `https://<DEIN-USERNAME>.github.io/<REPO-NAME>/Notion.ics`  
   Diese URL im Kalender als **Internetkalender / Abo** eintragen.

**Hinweis:** Jeder mit der URL kann die Termine abrufen – bei öffentlichem Repo/Pages ist die Kalenderdatei **öffentlich**. Token bleiben nur in GitHub Secrets.

## To-DO
* [ ] Make Quiz and Exams All Caps
