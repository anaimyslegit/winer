# Telegram Bot For Products

Файлы:

- `data/products.json` — основной каталог товаров
- `products.generated.js` — файл, который читает сайт
- `scripts/ProductStore.ps1` — функции чтения/сохранения каталога
- `telegram-product-bot.ps1` — Telegram-бот для добавления товаров
- `rebuild-products.ps1` — ручная пересборка `products.generated.js` из JSON

## 1. Создать бота

1. Напишите [@BotFather](https://t.me/BotFather)
2. Выполните `/newbot`
3. Сохраните токен

## 2. Задать переменные

В PowerShell:

```powershell
$env:TELEGRAM_BOT_TOKEN="YOUR_BOT_TOKEN"
$env:TELEGRAM_ADMIN_IDS="123456789"
```

`TELEGRAM_ADMIN_IDS` — необязательный список Telegram user id через запятую. Если не задан, бот примет команды от любого пользователя.

## 3. Запустить бота

```powershell
cd D:\winer
powershell -ExecutionPolicy Bypass -File .\telegram-product-bot.ps1
```

## 4. Команды в Telegram

```text
/start
/help
/categories
/list
/rebuild
/delete product-3
/add bike | Trek Marlin 7 | Алюмінієва рама, 29 дюймів, 18 швидкостей | 34990 | fa-bicycle
```

Формат `/add`:

```text
/add <category> | <name> | <description> | <price> | [icon]
```

Категории:

- `bike`
- `moto`
- `agro`
- `electro`

Если `icon` не указан, бот подставит его по категории.

## 5. Что происходит после добавления

1. Бот записывает товар в `data/products.json`
2. Автоматически пересобирает `products.generated.js`
3. После обновления страницы новый товар появляется на сайте
