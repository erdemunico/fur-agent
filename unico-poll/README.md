# Unico Poll (Slack Bolt + Node.js + SQLite)

Unico Poll, 3 asamali bir Slack anket uygulamasidir:
- Suggestion Phase
- Voting Phase
- Results & Run-off

## 1) Project Structure

```text
unico-poll/
  .env.example
  .gitignore
  package.json
  README.md
  data/
    .gitkeep
  src/
    index.js
    config/env.js
    db/migrations.js
    db/sqlite.js
    services/pollService.js
    services/scheduler.js
    slack/actions.js
    slack/blocks.js
    slack/commands.js
    utils/parser.js
    utils/time.js
```

## 2) package.json ve kurulum komutlari

```bash
cd "C:\Users\erdem\Downloads\fur agent\unico-poll"
npm install
copy .env.example .env
```

`.env` dosyasini doldurun:

```env
SLACK_BOT_TOKEN=xoxb-your-bot-token
SLACK_SIGNING_SECRET=your-signing-secret
SLACK_APP_TOKEN=xapp-your-app-token
PORT=3000
DATABASE_PATH=./data/unico-poll.db
DEFAULT_SUGGESTION_HOURS=48
DEFAULT_VOTING_HOURS=48
DEFAULT_RUNOFF_HOURS=24
FAST_TEST_MODE=false
FAST_TEST_MINUTES=5
ALLOWED_CREATOR_IDS=
SUGGESTION_RATE_LIMIT_COUNT=5
SUGGESTION_RATE_LIMIT_WINDOW_MINUTES=1
```

Calistirma:

```bash
npm run dev
```

Hizli test modu icin:

```env
FAST_TEST_MODE=true
FAST_TEST_MINUTES=5
```

Bu mod aktifken suggestion/voting/run-off sureleri saat yerine belirtilen dakika ile otomatik ayarlanir.

`ALLOWED_CREATOR_IDS` doldurulursa (ornegin `U123,U456`) sadece bu kullanicilar `/unico-poll` komutunu calistirabilir.

Oneri spam korumasi icin:
- `SUGGESTION_RATE_LIMIT_COUNT`: pencere icindeki max oneriler
- `SUGGESTION_RATE_LIMIT_WINDOW_MINUTES`: pencere suresi (dakika)

## 3) Slash command kullanimi

```text
/unico-poll Turnuva Ismi | 48h
```

- Oneri formati: `Oneri Ismi : PM Keyword ; Ekstra`
- Ankette sadece `Oneri Ismi` gorunur.
- `:` ve `;` sonrasi alanlar sadece log/PM takibi icin saklanir.

## 4) Veri modeli

- `polls`: anket temel bilgileri ve faz/sure alanlari
- `suggestions`: ham yazi + gorunen isim + pm keyword + extra
- `poll_shortlist`: ankete girecek secenekler (max 10)
- `votes_classic`: klasik oylar
- `votes_rating`: 1-5 puanlar
- `poll_events`: audit/log (genisletmeye hazir)

## 5) Slack app gerekli izinleri

- OAuth scopes: `commands`, `chat:write`, `chat:write.public`, `channels:history`, `groups:history`, `im:history`, `mpim:history`
- Slash command: `/unico-poll`
- Event subscription: `message.channels`, `message.groups`

## 6) GitHub'a ilk push adimlari

```bash
cd "C:\Users\erdem\Downloads\fur agent"
git add unico-poll
git commit -m "Add Unico Poll Slack Bolt app with SQLite state and run-off flow"
git branch -M main
git remote add origin https://github.com/<YOUR_USERNAME>/<YOUR_REPO>.git
git push -u origin main
```
