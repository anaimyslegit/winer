# Telegram Orders Setup

Files:

- `index.html` - checkout form and client-side order request
- `telegram-order-server.ps1` - local HTTP gateway that forwards orders to Telegram
- `data/orders/` - local backup copies of accepted orders

## 1. Prepare the bot

1. Open your bot in Telegram
2. Press `Start`
3. Make sure the bot can receive updates in `getUpdates`

## 2. Set environment variables

In PowerShell:

```powershell
$env:TELEGRAM_BOT_TOKEN="YOUR_BOT_TOKEN"
$env:TELEGRAM_ORDER_CHAT_ID="YOUR_CHAT_ID"
$env:NOVA_POSHTA_API_KEY="YOUR_NOVA_POSHTA_API_KEY"
```

`TELEGRAM_ORDER_CHAT_ID` is recommended. If you already use `TELEGRAM_ADMIN_IDS`, the order server can also reuse the first value from that list.
`NOVA_POSHTA_API_KEY` is used for branch lookup by city during checkout.

## 3. Start the local order server

```powershell
cd D:\winer
powershell -ExecutionPolicy Bypass -File .\telegram-order-server.ps1
```

Server endpoints:

- `http://127.0.0.1:8790/health`
- `http://127.0.0.1:8790/api/orders`
- `http://127.0.0.1:8790/api/nova-poshta/warehouses?city=Київ`

## 4. How it works

1. The customer opens the cart and clicks `Оформити`
2. The site can load Nova Poshta branches for the selected city from the local PowerShell server
3. The server forwards the order to Telegram
4. A backup copy is saved in `data/orders/`

## 5. If orders do not arrive

- Check that the server window is still running
- Check that the bot token is valid
- Press `Start` in Telegram again
- If chat id is not set, the server will try to resolve it from recent bot updates
