# Как перенести бота в облако (без банковской карты)

Бот на твоём ПК не может работать при выключенном компьютере. Нужно запустить его на сервере, который всегда включён. Ниже — бесплатный способ без карты.

## 1. GitHub (без карты)

1. Зайди на https://github.com и зарегистрируйся (нужна только почта).
2. Создай новый репозиторий: New repository → имя, например `school-bot` → **Public** (или Private) → Create repository.
3. Загрузи в него ТОЛЬКО эти файлы:
   - `school_bot.js`
   - `package.json`
   - `package-lock.json`
   - `render.yaml`
   - `.gitignore`
   
   **Не загружай** `node_modules`, `bot.log`, `.pptx`.

4. **Важно:** перед загрузкой открой `school_bot.js` и замени строку с токеном:

   Было:
   ```js
   const TOKEN = process.env.BOT_TOKEN || '9289379370:...';
   ```
   Стало:
   ```js
   const TOKEN = process.env.BOT_TOKEN;
   if (!TOKEN) throw new Error('Не задан BOT_TOKEN');
   ```

## 2. Render (без карты)

1. Зайди на https://render.com → **Sign up with GitHub** (карта не нужна).
2. Нажми **New + → Web Service** → подключи репозиторий `school-bot`.
3. Render сам увидит `render.yaml`. Нажми **Apply** / **Deploy Web Service**.
4. В настройках сервиса открой **Environment** → **Add Environment Variable**:
   - Key: `BOT_TOKEN`
   - Value: токен твоего бота от @BotFather
5. Дождись окончания деплоя. В логах должно быть:
   ```
   Webhook server started on port 3000
   Webhook URL: https://school-bot.onrender.com/webhook
   ```
6. Проверь в браузере `https://school-bot.onrender.com/` — должно быть `Bot is alive`.

## 3. Проверка

1. **Останови локального бота** (закрой чёрное окно `.bat`) — одновременно можно запускать только один экземпляр.
2. Напиши боту @schol_spg_bot → `/start`.
3. Напиши `напиши про энштейна`.
4. Выключи ПК — бот продолжит работать на Render.

## Важное

- Render бесплатно «усыпляет» сайт после ~15 минут тишины. Когда боту придёт сообщение, он проснётся — первый ответ может идти 10–30 секунд. Это нормально.
- Никогда не публикуй токен в открытом коде. Если он уже был в чате — отзови его (`/revoke` у @BotFather) и используй новый.
- Если видишь `409: Conflict` — значит где-то всё ещё запущена вторая копия бота (локальная или облачная). Оставь только одну.
