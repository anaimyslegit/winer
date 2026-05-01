param(
    [string]$BotToken = $env:TELEGRAM_BOT_TOKEN,
    [string]$ChatId = $env:TELEGRAM_ORDER_CHAT_ID,
    [string]$AdminIds = $env:TELEGRAM_ADMIN_IDS,
    [string]$NovaPoshtaApiKey = $env:NOVA_POSHTA_API_KEY,
    [int]$Port = 8790
)

function Get-FirstNonEmptyValue {
    param([string[]]$Values)

    foreach ($value in $Values) {
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim()
        }
    }

    return ""
}

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$BotToken = Get-FirstNonEmptyValue -Values @(
    $BotToken,
    [Environment]::GetEnvironmentVariable("TELEGRAM_BOT_TOKEN", "User")
)
$ChatId = Get-FirstNonEmptyValue -Values @(
    $ChatId,
    [Environment]::GetEnvironmentVariable("TELEGRAM_ORDER_CHAT_ID", "User")
)
$AdminIds = Get-FirstNonEmptyValue -Values @(
    $AdminIds,
    [Environment]::GetEnvironmentVariable("TELEGRAM_ADMIN_IDS", "User")
)
$NovaPoshtaApiKey = Get-FirstNonEmptyValue -Values @(
    $NovaPoshtaApiKey,
    [Environment]::GetEnvironmentVariable("NOVA_POSHTA_API_KEY", "User")
)

if ([string]::IsNullOrWhiteSpace($BotToken)) {
    throw "Set TELEGRAM_BOT_TOKEN before starting the order server."
}

$apiBase = "https://api.telegram.org/bot$BotToken"
$novaPoshtaApiBase = "https://api.novaposhta.ua/v2.0/json/"
$ordersRoot = Join-Path $PSScriptRoot "data\orders"
$orderStatusesPath = Join-Path $PSScriptRoot "data\order-statuses.json"
$warehouseCache = @{}
$cityLookupCache = @{}
$cityCatalogCache = $null
$cacheTtlMinutes = 20
$script:telegramUpdateOffset = 0
$orderStatuses = @{}

function Resolve-TelegramChatId {
    $adminChatId = ""
    if (-not [string]::IsNullOrWhiteSpace($AdminIds)) {
        $adminChatId = $AdminIds.Split(",", [System.StringSplitOptions]::RemoveEmptyEntries) |
            ForEach-Object { $_.Trim() } |
            Select-Object -First 1
    }

    $resolved = Get-FirstNonEmptyValue -Values @($ChatId, $adminChatId)
    if ($resolved) {
        return $resolved
    }

    try {
        $updates = Invoke-RestMethod -Uri "$apiBase/getUpdates?limit=20" -Method Get
        if ($updates.ok) {
            $chatIds = [System.Collections.Generic.List[string]]::new()

            foreach ($update in @($updates.result)) {
                if ($update.message -and $update.message.chat -and $update.message.chat.id) {
                    [void]$chatIds.Add([string]$update.message.chat.id)
                    continue
                }

                if ($update.callback_query -and $update.callback_query.message -and $update.callback_query.message.chat.id) {
                    [void]$chatIds.Add([string]$update.callback_query.message.chat.id)
                }
            }

            if ($chatIds.Count -gt 0) {
                return $chatIds[$chatIds.Count - 1]
            }
        }
    } catch {
        Write-Warning "Could not resolve chat id from Telegram updates: $($_.Exception.Message)"
    }

    return ""
}

$resolvedChatId = Resolve-TelegramChatId

if ([string]::IsNullOrWhiteSpace($resolvedChatId)) {
    throw "Could not resolve Telegram chat id. Set TELEGRAM_ORDER_CHAT_ID or open the bot in Telegram, press Start, and rerun the server."
}

function Write-CommonHeaders {
    param([System.Net.HttpListenerResponse]$Response)

    $Response.Headers["Access-Control-Allow-Origin"] = "*"
    $Response.Headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    $Response.Headers["Access-Control-Allow-Headers"] = "Content-Type"
}

function Write-JsonResponse {
    param(
        [Parameter(Mandatory = $true)][System.Net.HttpListenerContext]$Context,
        [Parameter(Mandatory = $true)][int]$StatusCode,
        [Parameter(Mandatory = $true)][object]$Payload
    )

    $response = $Context.Response
    $json = $Payload | ConvertTo-Json -Depth 12 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)

    Write-CommonHeaders -Response $response
    $response.StatusCode = $StatusCode
    $response.ContentType = "application/json; charset=utf-8"
    $response.ContentLength64 = $bytes.Length
    $response.OutputStream.Write($bytes, 0, $bytes.Length)
    $response.OutputStream.Close()
}

function Write-EmptyResponse {
    param(
        [Parameter(Mandatory = $true)][System.Net.HttpListenerContext]$Context,
        [Parameter(Mandatory = $true)][int]$StatusCode
    )

    $response = $Context.Response
    Write-CommonHeaders -Response $response
    $response.StatusCode = $StatusCode
    $response.ContentLength64 = 0
    $response.OutputStream.Close()
}

$script:telegramJsonSerializer = $null

function Ensure-TelegramJsonSerializer {
    if ($script:telegramJsonSerializer) {
        return $true
    }

    try {
        Add-Type -AssemblyName System.Web.Extensions -ErrorAction Stop
        $script:telegramJsonSerializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $script:telegramJsonSerializer.MaxJsonLength = [int]::MaxValue
        return $true
    } catch {
        return $false
    }
}

function ConvertTo-TelegramBotJson {
    param([object]$InputObject)

    if (Ensure-TelegramJsonSerializer) {
        return $script:telegramJsonSerializer.Serialize($InputObject)
    }

    return ($InputObject | ConvertTo-Json -Compress -Depth 20)
}

function Get-TelegramApiErrorBody {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)

    $exception = $ErrorRecord.Exception
    while ($exception -and -not $exception.Response -and $exception.InnerException) {
        $exception = $exception.InnerException
    }

    if (-not $exception -or -not $exception.Response) {
        return ""
    }

    try {
        $stream = $exception.Response.GetResponseStream()
        if (-not $stream) {
            return ""
        }

        $reader = New-Object System.IO.StreamReader($stream)
        try {
            return $reader.ReadToEnd()
        } finally {
            $reader.Dispose()
        }
    } catch {
        return ""
    }
}

function Get-RequestEncoding {
    param([System.Net.HttpListenerRequest]$Request)

    $utf8 = [System.Text.UTF8Encoding]::new($false, $false)
    $charset = ""

    if ($Request -and $Request.ContentType -match "charset\s*=\s*([^;]+)") {
        $charset = ($Matches[1] -replace '"', "").Trim()
    }

    if ([string]::IsNullOrWhiteSpace($charset)) {
        return $utf8
    }

    try {
        return [System.Text.Encoding]::GetEncoding($charset)
    } catch {
        return $utf8
    }
}

function Read-RequestBody {
    param([System.Net.HttpListenerRequest]$Request)

    $encoding = Get-RequestEncoding -Request $Request
    $reader = [System.IO.StreamReader]::new($Request.InputStream, $encoding, $true)
    try {
        return $reader.ReadToEnd()
    } finally {
        $reader.Dispose()
    }
}

function Get-QueryParameters {
    param([string]$QueryString)

    $result = @{}
    $normalizedQuery = [string]$QueryString

    if ([string]::IsNullOrWhiteSpace($normalizedQuery)) {
        return $result
    }

    $normalizedQuery = $normalizedQuery.TrimStart("?")
    foreach ($pair in ($normalizedQuery -split "&")) {
        if ([string]::IsNullOrWhiteSpace($pair)) {
            continue
        }

        $segments = $pair -split "=", 2
        $rawKey = if ($segments.Count -ge 1) { $segments[0] } else { "" }
        $rawValue = if ($segments.Count -ge 2) { $segments[1] } else { "" }

        $decodedKey = [System.Uri]::UnescapeDataString(($rawKey -replace "\+", " "))
        $decodedValue = [System.Uri]::UnescapeDataString(($rawValue -replace "\+", " "))

        if ($decodedKey) {
            $result[$decodedKey] = $decodedValue
        }
    }

    return $result
}

function New-OrderId {
    return "ORD-" + (Get-Date -Format "yyyyMMdd-HHmmss")
}

function Format-PriceValue {
    param([decimal]$Value)

    return ("{0:N0} UAH" -f $Value).Replace(",", " ")
}

function Get-DeliveryLabel {
    param([string]$Value)

    $labels = @{
        "nova-poshta-warehouse" = "Nova Poshta branch"
        "nova-poshta-courier"   = "Nova Poshta courier"
        "store-pickup"          = "Store pickup"
    }

    if ($labels.ContainsKey($Value)) {
        return $labels[$Value]
    }

    return $Value
}

function Get-PaymentLabel {
    param([string]$Value)

    $labels = @{
        "card"             = "Card payment"
        "cash-on-delivery" = "Cash on delivery"
        "bank-transfer"    = "Bank transfer"
    }

    if ($labels.ContainsKey($Value)) {
        return $labels[$Value]
    }

    return $Value
}

function Require-OrderField {
    param(
        [string]$Value,
        [string]$FieldName
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "Missing required field: $FieldName"
    }

    return $Value.Trim()
}

function Save-OrderBackup {
    param(
        [string]$OrderId,
        [object]$Order
    )

    if (-not (Test-Path -LiteralPath $ordersRoot)) {
        New-Item -ItemType Directory -Path $ordersRoot | Out-Null
    }

    $path = Join-Path $ordersRoot "$OrderId.json"
    $json = $Order | ConvertTo-Json -Depth 12
    Set-Content -LiteralPath $path -Value $json -Encoding UTF8
}

function Ensure-DataDirectory {
    $dataDir = Split-Path -Path $orderStatusesPath -Parent
    if (-not (Test-Path -LiteralPath $dataDir)) {
        New-Item -ItemType Directory -Path $dataDir | Out-Null
    }
}

function Initialize-OrderStatuses {
    Ensure-DataDirectory

    if (-not (Test-Path -LiteralPath $orderStatusesPath)) {
        $orderStatuses = @{}
        return
    }

    try {
        $raw = Get-Content -LiteralPath $orderStatusesPath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) {
            $orderStatuses = @{}
            return
        }

        $parsed = $raw | ConvertFrom-Json
        $loaded = @{}
        foreach ($prop in $parsed.PSObject.Properties) {
            $loaded[$prop.Name] = $prop.Value
        }
        $orderStatuses = $loaded
    } catch {
        Write-Warning "Could not load order statuses: $($_.Exception.Message)"
        $orderStatuses = @{}
    }
}

function Save-OrderStatuses {
    Ensure-DataDirectory
    $json = $orderStatuses | ConvertTo-Json -Depth 8
    Set-Content -LiteralPath $orderStatusesPath -Value $json -Encoding UTF8
}

function Get-OrderStatusLabel {
    param([string]$StatusCode)

    $labels = @{
        "payment_pending"   = "Payment pending confirmation"
        "payment_confirmed" = "Payment confirmed"
        "payment_rejected"  = "Payment rejected"
    }

    if ($labels.ContainsKey($StatusCode)) {
        return $labels[$StatusCode]
    }

    return "Processing"
}

function Set-OrderStatus {
    param(
        [Parameter(Mandatory = $true)][string]$OrderId,
        [Parameter(Mandatory = $true)][string]$StatusCode,
        [string]$UpdatedBy = "system"
    )

    $orderStatuses[$OrderId] = [pscustomobject]@{
        code      = $StatusCode
        label     = Get-OrderStatusLabel -StatusCode $StatusCode
        updatedAt = (Get-Date).ToString("o")
        updatedBy = $UpdatedBy
    }

    Save-OrderStatuses
}

function Get-OrderStatus {
    param([Parameter(Mandatory = $true)][string]$OrderId)

    if ($orderStatuses.ContainsKey($OrderId)) {
        return $orderStatuses[$OrderId]
    }

    return [pscustomobject]@{
        code      = "payment_pending"
        label     = Get-OrderStatusLabel -StatusCode "payment_pending"
        updatedAt = ""
        updatedBy = ""
    }
}

function Get-PaymentInlineKeyboard {
    param([Parameter(Mandatory = $true)][string]$OrderId)

    return @{
        inline_keyboard = @(
            @(
                @{
                    text          = [string]::Concat([char]0x2705, " Підтвердити оплату")
                    callback_data = "pay:confirm:$OrderId"
                },
                @{
                    text          = [string]::Concat([char]0x274C, " Відхилити")
                    callback_data = "pay:reject:$OrderId"
                }
            )
        )
    }
}

function Build-OrderMessage {
    param(
        [string]$OrderId,
        [object]$Order
    )

    $customer = $Order.customer
    $totals = $Order.totals
    $items = @($Order.items)

    $formattedTotal = if ($totals.formattedTotal) {
        [string]$totals.formattedTotal
    } else {
        Format-PriceValue -Value ([decimal]$totals.totalPrice)
    }

    $lines = [System.Collections.Generic.List[string]]::new()
    [void]$lines.Add("New order: $OrderId")
    [void]$lines.Add("Time: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))")
    [void]$lines.Add("")
    [void]$lines.Add("Customer")
    [void]$lines.Add("Name: $($customer.name)")
    [void]$lines.Add("Last name: $($customer.lastName)")
    [void]$lines.Add("Phone: $($customer.phone)")
    [void]$lines.Add("Email: $($customer.email)")
    [void]$lines.Add("Region: $($customer.region)")
    [void]$lines.Add("City: $($customer.city)")
    [void]$lines.Add("Delivery: $(Get-DeliveryLabel -Value ([string]$customer.delivery))")

    if (-not [string]::IsNullOrWhiteSpace([string]$customer.warehouseLabel)) {
        [void]$lines.Add("Branch: $($customer.warehouseLabel)")
    }

    [void]$lines.Add("Payment: $(Get-PaymentLabel -Value ([string]$customer.payment))")

    if (-not [string]::IsNullOrWhiteSpace([string]$customer.comment)) {
        [void]$lines.Add("Comment: $($customer.comment)")
    }

    [void]$lines.Add("")
    [void]$lines.Add("Items")

    foreach ($item in $items) {
        $subtotal = if ($item.formattedSubtotal) {
            [string]$item.formattedSubtotal
        } else {
            Format-PriceValue -Value ([decimal]$item.subtotal)
        }

        [void]$lines.Add("- $($item.name) | qty: $($item.quantity) | $subtotal")
    }

    [void]$lines.Add("")
    [void]$lines.Add("Quantity: $($totals.quantity)")
    [void]$lines.Add("Total: $formattedTotal")

    return ($lines -join "`n")
}

function Send-TelegramOrder {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $true)][string]$OrderId
    )

    $sendPayload = @{
        chat_id = $resolvedChatId
        text    = $Message
    }

    $sendBody = [System.Text.Encoding]::UTF8.GetBytes((ConvertTo-TelegramBotJson $sendPayload))

    try {
        $sent = Invoke-RestMethod -Uri "$apiBase/sendMessage" -Method Post -ContentType "application/json; charset=utf-8" -Body $sendBody
        if (-not $sent.ok) {
            throw "Telegram sendMessage returned ok=false."
        }

        $messageId = [int64]$sent.result.message_id
        $editPayload = @{
            chat_id      = $resolvedChatId
            message_id   = $messageId
            reply_markup = Get-PaymentInlineKeyboard -OrderId $OrderId
        }

        $editBody = [System.Text.Encoding]::UTF8.GetBytes((ConvertTo-TelegramBotJson $editPayload))
        $edited = Invoke-RestMethod -Uri "$apiBase/editMessageReplyMarkup" -Method Post -ContentType "application/json; charset=utf-8" -Body $editBody

        if (-not $edited.ok) {
            Write-Warning "Telegram editMessageReplyMarkup returned ok=false (buttons may be missing)."
        }
    } catch {
        $rawError = Get-TelegramApiErrorBody -ErrorRecord $_

        if (-not [string]::IsNullOrWhiteSpace($rawError)) {
            throw "Telegram API error: $rawError"
        }

        throw $_
    }
}

function Sync-TelegramPaymentUpdates {
    try {
        $query = "limit=50&allowed_updates=%5B%22callback_query%22%5D"
        if ($script:telegramUpdateOffset -gt 0) {
            $query += "&offset=$($script:telegramUpdateOffset)"
        }

        $updates = Invoke-RestMethod -Uri "$apiBase/getUpdates?$query" -Method Get
        if (-not $updates.ok) {
            return
        }

        foreach ($update in @($updates.result)) {
            if ($update.update_id -and ([int64]$update.update_id -ge $script:telegramUpdateOffset)) {
                $script:telegramUpdateOffset = [int64]$update.update_id + 1
            }

            $callback = $update.callback_query
            if (-not $callback) {
                continue
            }

            $callbackData = [string]$callback.data
            $callbackId = [string]$callback.id
            $handled = $false
            $alertText = "Unknown action."

            if ($callbackData -match "^pay:(confirm|reject):(ORD-\d{8}-\d{6})$") {
                $action = $Matches[1]
                $orderId = $Matches[2]
                $statusCode = if ($action -eq "confirm") { "payment_confirmed" } else { "payment_rejected" }
                $updatedBy = if ($callback.from.username) { "@$($callback.from.username)" } else { "telegram-admin" }
                Set-OrderStatus -OrderId $orderId -StatusCode $statusCode -UpdatedBy $updatedBy
                $handled = $true
                $alertText = "Order #${orderId}: $(Get-OrderStatusLabel -StatusCode $statusCode)"
            }

            if (-not [string]::IsNullOrWhiteSpace($callbackId)) {
                $answerPayload = @{
                    callback_query_id = $callbackId
                    text              = $alertText
                    show_alert        = $false
                }
                $answerBody = [System.Text.Encoding]::UTF8.GetBytes((ConvertTo-TelegramBotJson $answerPayload))
                [void](Invoke-RestMethod -Uri "$apiBase/answerCallbackQuery" -Method Post -ContentType "application/json; charset=utf-8" -Body $answerBody)
            }

            if ($handled -and $callback.message -and $callback.message.chat -and $callback.message.message_id) {
                $chatIdForEdit = [string]$callback.message.chat.id
                $messageIdForEdit = [string]$callback.message.message_id
                $clearKeyboardPayload = @{
                    chat_id      = $chatIdForEdit
                    message_id   = $messageIdForEdit
                    reply_markup = @{ inline_keyboard = @() }
                }
                $clearKeyboardBody = [System.Text.Encoding]::UTF8.GetBytes((ConvertTo-TelegramBotJson $clearKeyboardPayload))
                try {
                    [void](Invoke-RestMethod -Uri "$apiBase/editMessageReplyMarkup" -Method Post -ContentType "application/json; charset=utf-8" -Body $clearKeyboardBody)
                } catch {
                    # Ignore if message was already updated.
                }
            }
        }
    } catch {
        Write-Warning "Could not sync Telegram callback updates: $($_.Exception.Message)"
    }
}

function ConvertTo-ValidatedOrder {
    param([string]$RawBody)

    if ([string]::IsNullOrWhiteSpace($RawBody)) {
        throw "Request body is empty."
    }

    $order = $RawBody | ConvertFrom-Json

    if (-not $order.customer) {
        throw "Customer data is missing."
    }

    if (-not $order.items -or @($order.items).Count -eq 0) {
        throw "Order must contain at least one item."
    }

    [void](Require-OrderField -Value ([string]$order.customer.name) -FieldName "customer.name")
    [void](Require-OrderField -Value ([string]$order.customer.lastName) -FieldName "customer.lastName")
    [void](Require-OrderField -Value ([string]$order.customer.phone) -FieldName "customer.phone")
    [void](Require-OrderField -Value ([string]$order.customer.email) -FieldName "customer.email")
    [void](Require-OrderField -Value ([string]$order.customer.region) -FieldName "customer.region")
    [void](Require-OrderField -Value ([string]$order.customer.city) -FieldName "customer.city")
    [void](Require-OrderField -Value ([string]$order.customer.delivery) -FieldName "customer.delivery")
    [void](Require-OrderField -Value ([string]$order.customer.payment) -FieldName "customer.payment")

    if ([string]$order.customer.delivery -eq "nova-poshta-warehouse") {
        [void](Require-OrderField -Value ([string]$order.customer.warehouseRef) -FieldName "customer.warehouseRef")
        [void](Require-OrderField -Value ([string]$order.customer.warehouseLabel) -FieldName "customer.warehouseLabel")
    }

    foreach ($item in @($order.items)) {
        [void](Require-OrderField -Value ([string]$item.name) -FieldName "item.name")

        if (([int]$item.quantity) -le 0) {
            throw "Item quantity must be greater than 0."
        }
    }

    return $order
}

function Get-NovaPoshtaErrorMessage {
    param([object]$Response)

    $parts = @()

    if ($Response.errors) {
        $parts += @($Response.errors)
    }

    if ($Response.info) {
        if ($Response.info -is [System.Collections.IDictionary]) {
            foreach ($key in $Response.info.Keys) {
                $parts += ("{0}: {1}" -f $key, $Response.info[$key])
            }
        } else {
            $parts += @($Response.info)
        }
    }

    if ($parts.Count -eq 0) {
        return "Nova Poshta API request failed."
    }

    return ($parts -join " | ")
}

function Invoke-NovaPoshtaApi {
    param(
        [Parameter(Mandatory = $true)][string]$ModelName,
        [Parameter(Mandatory = $true)][string]$CalledMethod,
        [Parameter(Mandatory = $true)][hashtable]$MethodProperties
    )

    if ([string]::IsNullOrWhiteSpace($NovaPoshtaApiKey)) {
        throw "Nova Poshta API key is not configured."
    }

    $payload = @{
        apiKey           = $NovaPoshtaApiKey
        modelName        = $ModelName
        calledMethod     = $CalledMethod
        methodProperties = $MethodProperties
    }
    $requestBody = [System.Text.Encoding]::UTF8.GetBytes(($payload | ConvertTo-Json -Depth 8))

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            $response = Invoke-RestMethod -Uri $novaPoshtaApiBase -Method Post -ContentType "application/json; charset=utf-8" -Body $requestBody
        } catch {
            $httpResponse = $_.Exception.Response
            $rawApiError = ""

            if ($httpResponse) {
                try {
                    $stream = $httpResponse.GetResponseStream()
                    if ($stream) {
                        $reader = New-Object System.IO.StreamReader($stream)
                        $rawApiError = $reader.ReadToEnd()
                        $reader.Dispose()
                    }
                } catch {
                    $rawApiError = ""
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($rawApiError)) {
                try {
                    $parsedError = $rawApiError | ConvertFrom-Json
                    $normalizedError = Get-NovaPoshtaErrorMessage -Response $parsedError
                    if (-not [string]::IsNullOrWhiteSpace($normalizedError)) {
                        throw "Nova Poshta API error: $normalizedError"
                    }
                } catch {
                    throw "Nova Poshta API error: $rawApiError"
                }
            }

            throw $_
        }

        if ($response.success) {
            return $response
        }

        $errorMessage = Get-NovaPoshtaErrorMessage -Response $response
        if ($errorMessage -match "To many requests" -and $attempt -lt 3) {
            Start-Sleep -Milliseconds 700
            continue
        }

        throw $errorMessage
    }

    throw "Nova Poshta API request failed after retries."
}

function Normalize-LookupText {
    param([string]$Value)

    $text = [string]$Value

    return [string]::Join(
        " ",
        $text.Trim().ToLowerInvariant().Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)
    )
}

function Get-ObjectStringProperty {
    param(
        [Parameter(Mandatory = $true)][object]$Object,
        [Parameter(Mandatory = $true)][string]$PropertyName
    )

    $property = $Object.PSObject.Properties[$PropertyName]
    if ($property) {
        return [string]$property.Value
    }

    return ""
}

function Get-NovaPoshtaCityCatalog {
    if ($cityCatalogCache -and $cityCatalogCache.ExpiresAt -gt (Get-Date)) {
        return $cityCatalogCache.Items
    }

    $response = Invoke-NovaPoshtaApi `
        -ModelName "Address" `
        -CalledMethod "getCities" `
        -MethodProperties @{}

    $items = @($response.data)
    $cityCatalogCache = @{
        ExpiresAt = (Get-Date).AddHours(12)
        Items     = $items
    }

    return $items
}

function Test-ContainsCyrillic {
    param([string]$Value)

    return [regex]::IsMatch([string]$Value, "[\u0400-\u052F]")
}

function Get-TextFromCodePoints {
    param([int[]]$CodePoints)

    return [string]::Join("", ($CodePoints | ForEach-Object { [char]$_ }))
}

function Get-NovaPoshtaCityLookupName {
    param([string]$City)

    $trimmed = ([string]$City).Trim()
    $key = $trimmed.ToLowerInvariant()

    switch -Regex ($key) {
        "^(kyiv|kiev)$" { return Get-TextFromCodePoints @(0x041A, 0x0438, 0x0457, 0x0432) }
        "^(lviv|lvov)$" { return Get-TextFromCodePoints @(0x041B, 0x044C, 0x0432, 0x0456, 0x0432) }
        "^(odesa|odessa)$" { return Get-TextFromCodePoints @(0x041E, 0x0434, 0x0435, 0x0441, 0x0430) }
        "^(kharkiv|harkiv|kharkov)$" { return Get-TextFromCodePoints @(0x0425, 0x0430, 0x0440, 0x043A, 0x0456, 0x0432) }
        "^(dnipro|dnepr)$" { return Get-TextFromCodePoints @(0x0414, 0x043D, 0x0456, 0x043F, 0x0440, 0x043E) }
        "^(zaporizhzhia|zaporizhia|zaporozhye)$" { return Get-TextFromCodePoints @(0x0417, 0x0430, 0x043F, 0x043E, 0x0440, 0x0456, 0x0436, 0x0436, 0x044F) }
        "^(vinnytsia|vinnitsa)$" { return Get-TextFromCodePoints @(0x0412, 0x0456, 0x043D, 0x043D, 0x0438, 0x0446, 0x044F) }
        "^(chernivtsi)$" { return Get-TextFromCodePoints @(0x0427, 0x0435, 0x0440, 0x043D, 0x0456, 0x0432, 0x0446, 0x0456) }
        default { return $trimmed }
    }
}

function Resolve-NovaPoshtaCityFromSearch {
    param(
        [Parameter(Mandatory = $true)][string]$City,
        [string]$Region = ""
    )

    $needle = Normalize-LookupText -Value $City
    $regionNeedle = Normalize-LookupText -Value $Region
    if ($needle.Length -lt 2) {
        throw "City name is too short."
    }

    $response = Invoke-NovaPoshtaApi `
        -ModelName "Address" `
        -CalledMethod "searchSettlements" `
        -MethodProperties @{
            CityName = $City
            Limit    = "20"
        }

    $responseGroups = @($response.data)
    $cityTypeCode = [string]([char]0x043C) + "."

    $candidates = for ($groupIndex = 0; $groupIndex -lt $responseGroups.Count; $groupIndex++) {
        $group = $responseGroups[$groupIndex]
        $addresses = @($group.Addresses)

        for ($addressIndex = 0; $addressIndex -lt $addresses.Count; $addressIndex++) {
            $item = $addresses[$addressIndex]
            $present = Normalize-LookupText -Value (Get-ObjectStringProperty -Object $item -PropertyName "Present")
            $mainDescription = Normalize-LookupText -Value (Get-ObjectStringProperty -Object $item -PropertyName "MainDescription")
            $area = Normalize-LookupText -Value (Get-ObjectStringProperty -Object $item -PropertyName "Area")
            $district = Normalize-LookupText -Value (Get-ObjectStringProperty -Object $item -PropertyName "Region")
            $settlementTypeCode = Normalize-LookupText -Value (Get-ObjectStringProperty -Object $item -PropertyName "SettlementTypeCode")
            $warehouseCount = 0
            [void][int]::TryParse((Get-ObjectStringProperty -Object $item -PropertyName "Warehouses"), [ref]$warehouseCount)

            $score = [math]::Max(0, 140 - ($groupIndex * 20) - ($addressIndex * 5))

            if ($mainDescription -eq $needle) {
                $score = 320
            } elseif ($present -eq $needle) {
                $score = 300
            } elseif ($mainDescription.StartsWith($needle) -or $present.StartsWith($needle)) {
                $score = [math]::Max($score, 210)
            } elseif ($mainDescription.Contains($needle) -or $present.Contains($needle)) {
                $score = [math]::Max($score, 120)
            }

            if ($settlementTypeCode -eq "м.") {
                $score += 30
            }

            if ($settlementTypeCode -eq $cityTypeCode) {
                $score += 30
            }

            if ($regionNeedle.Length -ge 2) {
                if ((($area.Length -ge 2) -and ($area.Contains($regionNeedle) -or $regionNeedle.Contains($area))) -or (($district.Length -ge 2) -and $district.Contains($regionNeedle)) -or $present.Contains($regionNeedle)) {
                    $score += 70
                } else {
                    $score -= 30
                }
            }

            if ($warehouseCount -gt 0) {
                $score += [math]::Min(100, [int][math]::Ceiling([math]::Log10($warehouseCount + 1) * 35))
            } else {
                $score -= 80
            }

            [pscustomobject]@{
                Ref         = Get-ObjectStringProperty -Object $item -PropertyName "DeliveryCity"
                Description = Get-FirstNonEmptyValue -Values @(
                    (Get-ObjectStringProperty -Object $item -PropertyName "MainDescription"),
                    (Get-ObjectStringProperty -Object $item -PropertyName "Present")
                )
                Score       = $score
            }
        }
    }

    $resolved = @(
        $candidates |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.Ref) } |
            Sort-Object @{ Expression = "Score"; Descending = $true }, @{ Expression = "Description"; Descending = $false } |
            Select-Object -First 1
    )

    if ($resolved.Count -eq 0) {
        $fallbackItems = @(
            for ($groupIndex = 0; $groupIndex -lt $responseGroups.Count; $groupIndex++) {
                $group = $responseGroups[$groupIndex]
                $addresses = @($group.Addresses)

                for ($addressIndex = 0; $addressIndex -lt $addresses.Count; $addressIndex++) {
                    $item = $addresses[$addressIndex]
                    $deliveryCityRef = Get-ObjectStringProperty -Object $item -PropertyName "DeliveryCity"
                    if ([string]::IsNullOrWhiteSpace($deliveryCityRef)) {
                        continue
                    }

                    [pscustomobject]@{
                        Ref         = $deliveryCityRef
                        Description = Get-FirstNonEmptyValue -Values @(
                            (Get-ObjectStringProperty -Object $item -PropertyName "MainDescription"),
                            (Get-ObjectStringProperty -Object $item -PropertyName "Present")
                        )
                    }
                }
            }
        )
        $fallbackCandidate = @($fallbackItems | Select-Object -First 1)

        if ($fallbackCandidate.Count -gt 0) {
            return $fallbackCandidate[0]
        }

        throw "Could not resolve Nova Poshta city reference for '$City'."
    }

    return $resolved[0]
}

function Resolve-NovaPoshtaCityFromCatalog {
    param(
        [Parameter(Mandatory = $true)][string]$City,
        [string]$Region = ""
    )

    $needle = Normalize-LookupText -Value $City
    $regionNeedle = Normalize-LookupText -Value $Region
    if ($needle.Length -lt 2) {
        throw "City name is too short."
    }

    $catalog = Get-NovaPoshtaCityCatalog
    $candidates = foreach ($item in $catalog) {
        $descriptionUa = Normalize-LookupText -Value (Get-ObjectStringProperty -Object $item -PropertyName "Description")
        $descriptionRu = Normalize-LookupText -Value (Get-ObjectStringProperty -Object $item -PropertyName "DescriptionRu")
        $descriptionTranslit = Normalize-LookupText -Value (Get-ObjectStringProperty -Object $item -PropertyName "DescriptionTranslit")
        $typeUa = Normalize-LookupText -Value (Get-ObjectStringProperty -Object $item -PropertyName "SettlementTypeDescription")
        $typeRu = Normalize-LookupText -Value (Get-ObjectStringProperty -Object $item -PropertyName "SettlementTypeDescriptionRu")
        $areaUa = Normalize-LookupText -Value (Get-ObjectStringProperty -Object $item -PropertyName "AreaDescription")
        $areaRu = Normalize-LookupText -Value (Get-ObjectStringProperty -Object $item -PropertyName "AreaDescriptionRu")
        $hasWarehouse = (Get-ObjectStringProperty -Object $item -PropertyName "Warehouse") -eq "1"

        $score = -1

        if ($descriptionUa -eq $needle -or $descriptionRu -eq $needle -or $descriptionTranslit -eq $needle) {
            $score = 300
        } elseif ($descriptionUa.StartsWith($needle) -or $descriptionRu.StartsWith($needle) -or $descriptionTranslit.StartsWith($needle)) {
            $score = 200
        } elseif ($descriptionUa.Contains($needle) -or $descriptionRu.Contains($needle) -or $descriptionTranslit.Contains($needle)) {
            $score = 100
        }

        if ($score -lt 0) {
            continue
        }

        if ($typeUa -eq "місто" -or $typeRu -eq "город") {
            $score += 30
        }

        if ($regionNeedle.Length -ge 2) {
            if ((($areaUa.Length -ge 2) -and ($areaUa.Contains($regionNeedle) -or $regionNeedle.Contains($areaUa))) -or (($areaRu.Length -ge 2) -and ($areaRu.Contains($regionNeedle) -or $regionNeedle.Contains($areaRu)))) {
                $score += 70
            } else {
                $score -= 30
            }
        }

        if ($hasWarehouse) {
            $score += 15
        }

        [pscustomobject]@{
            Ref         = Get-ObjectStringProperty -Object $item -PropertyName "Ref"
            Description = Get-ObjectStringProperty -Object $item -PropertyName "Description"
            Score       = $score
        }
    }

    $resolved = @(
        $candidates |
            Sort-Object @{ Expression = "Score"; Descending = $true }, @{ Expression = "Description"; Descending = $false } |
            Select-Object -First 1
    )
    if ($resolved.Count -eq 0) {
        throw "Could not resolve Nova Poshta city reference for '$City'."
    }

    return $resolved[0]
}

function Resolve-NovaPoshtaCity {
    param(
        [Parameter(Mandatory = $true)][string]$City,
        [string]$Region = ""
    )

    $needle = Normalize-LookupText -Value $City
    $regionNeedle = Normalize-LookupText -Value $Region
    if ($needle.Length -lt 2) {
        throw "City name is too short."
    }

    $cacheKey = "$needle|$regionNeedle"
    $cached = $cityLookupCache[$cacheKey]

    if ($cached -and $cached.ExpiresAt -gt (Get-Date)) {
        return $cached.Payload
    }

    $resolved = $null
    $lookupCandidates = @(
        [string]$City,
        (Get-NovaPoshtaCityLookupName -City $City)
    ) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique

    foreach ($candidateCity in $lookupCandidates) {
        try {
            $resolved = Resolve-NovaPoshtaCityFromSearch -City $candidateCity -Region $Region
            if ($resolved) {
                break
            }
        } catch {
            $resolved = $null
        }
    }

    if (-not $resolved) {
        foreach ($candidateCity in $lookupCandidates) {
            try {
                $resolved = Resolve-NovaPoshtaCityFromCatalog -City $candidateCity -Region $Region
                if ($resolved) {
                    break
                }
            } catch {
                $resolved = $null
            }
        }
    }

    if (-not $resolved) {
        throw "Could not resolve Nova Poshta city reference for '$City'. Check city spelling (for example: Kyiv, Lviv, Odesa) and try again."
    }

    $cityLookupCache[$cacheKey] = @{
        ExpiresAt = (Get-Date).AddMinutes($cacheTtlMinutes)
        Payload   = $resolved
    }

    return $resolved
}

function Invoke-NovaPoshtaWarehouseLookup {
    param(
        [Parameter(Mandatory = $true)][string]$CityRef,
        [int]$Limit = 50
    )

    return Invoke-NovaPoshtaApi `
        -ModelName "AddressGeneral" `
        -CalledMethod "getWarehouses" `
        -MethodProperties @{
            CityRef  = $CityRef
            Language = "UA"
            Limit    = [string]$Limit
            Page     = "1"
        }
}

function Get-NovaPoshtaWarehousesByCity {
    param(
        [Parameter(Mandatory = $true)][string]$City,
        [string]$Region = "",
        [int]$Limit = 50
    )

    $normalizedCity = Get-NovaPoshtaCityLookupName -City $City
    $normalizedRegion = ([string]$Region).Trim()
    if ($normalizedCity.Length -lt 2) {
        throw "City name is too short."
    }

    $cacheKey = "$($normalizedCity.ToLowerInvariant())|$($normalizedRegion.ToLowerInvariant())"
    $cached = $warehouseCache[$cacheKey]

    if ($cached -and $cached.ExpiresAt -gt (Get-Date)) {
        return $cached.Payload
    }

    $resolvedCity = Resolve-NovaPoshtaCity -City $normalizedCity -Region $normalizedRegion
    $response = Invoke-NovaPoshtaWarehouseLookup -CityRef $resolvedCity.Ref -Limit $Limit

    if (@($response.data).Count -eq 0 -and (Test-ContainsCyrillic -Value $normalizedCity)) {
        try {
            $catalogCity = Resolve-NovaPoshtaCityFromCatalog -City $normalizedCity -Region $normalizedRegion
            if ($catalogCity.Ref -and $catalogCity.Ref -ne $resolvedCity.Ref) {
                $fallbackResponse = Invoke-NovaPoshtaWarehouseLookup -CityRef $catalogCity.Ref -Limit $Limit
                if (@($fallbackResponse.data).Count -gt 0) {
                    $resolvedCity = $catalogCity
                    $response = $fallbackResponse
                    $cityLookupCache[$cacheKey] = @{
                        ExpiresAt = (Get-Date).AddMinutes($cacheTtlMinutes)
                        Payload   = $resolvedCity
                    }
                }
            }
        } catch {
            Write-Warning "Nova Poshta catalog fallback failed for '$normalizedCity': $($_.Exception.Message)"
        }
    }

    $warehouses = @(
        @($response.data) |
            Select-Object -First $Limit |
            ForEach-Object {
                $shortAddress = if ($_.ShortAddress) {
                    [string]$_.ShortAddress
                } else {
                    [string]$_.Description
                }

                [pscustomobject]@{
                    ref          = [string]$_.Ref
                    number       = [string]$_.Number
                    label        = [string]$_.Description
                    shortAddress = $shortAddress
                    phone        = [string]$_.Phone
                    category     = [string]$_.CategoryOfWarehouse
                }
            }
    )

    $totalCount = if ($response.info.totalCount) {
        [int]$response.info.totalCount
    } else {
        $warehouses.Count
    }

    $payload = [pscustomobject]@{
        city        = $resolvedCity.Description
        totalCount  = $totalCount
        warehouses  = $warehouses
    }

    $warehouseCache[$cacheKey] = @{
        ExpiresAt = (Get-Date).AddMinutes($cacheTtlMinutes)
        Payload   = $payload
    }

    return $payload
}

$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("http://127.0.0.1:$Port/")
Initialize-OrderStatuses
$listener.Start()

Write-Host "Telegram order server is running."
Write-Host "Endpoint: http://127.0.0.1:$Port/api/orders"
Write-Host "Health:   http://127.0.0.1:$Port/health"
Write-Host "NP API:   http://127.0.0.1:$Port/api/nova-poshta/warehouses?city=Kyiv"
Write-Host "Chat ID:  $resolvedChatId"

try {
    while ($listener.IsListening) {
        $asyncResult = $listener.BeginGetContext($null, $null)
        while (-not $asyncResult.AsyncWaitHandle.WaitOne(2500)) {
            Sync-TelegramPaymentUpdates
        }

        $context = $listener.EndGetContext($asyncResult)

        try {
            $request = $context.Request
            $path = $request.Url.AbsolutePath.TrimEnd("/")

            if ($request.HttpMethod -eq "OPTIONS") {
                Write-EmptyResponse -Context $context -StatusCode 204
                continue
            }

            if ($request.HttpMethod -eq "GET" -and ($path -eq "/health" -or $path -eq "")) {
                Write-JsonResponse -Context $context -StatusCode 200 -Payload @{
                    ok                  = $true
                    service             = "telegram-order-server"
                    chatId              = $resolvedChatId
                    novaPoshtaConfigured = (-not [string]::IsNullOrWhiteSpace($NovaPoshtaApiKey))
                    paymentConfirmations = $true
                }
                continue
            }

            if ($request.HttpMethod -eq "GET" -and $path -eq "/api/nova-poshta/warehouses") {
                $query = Get-QueryParameters -QueryString $request.Url.Query
                $city = [string]$query["city"]
                $region = [string]$query["region"]
                $limitRaw = [string]$query["limit"]
                $limit = 50

                if ($limitRaw -and [int]::TryParse($limitRaw, [ref]$limit)) {
                    $limit = [Math]::Min([Math]::Max($limit, 1), 100)
                } else {
                    $limit = 50
                }

                $payload = Get-NovaPoshtaWarehousesByCity -City $city -Region $region -Limit $limit
                Write-JsonResponse -Context $context -StatusCode 200 -Payload @{
                    ok         = $true
                    city       = $payload.city
                    totalCount = $payload.totalCount
                    warehouses = $payload.warehouses
                }
                continue
            }

            if ($request.HttpMethod -eq "GET" -and $path -eq "/api/orders/status") {
                Sync-TelegramPaymentUpdates
                $query = Get-QueryParameters -QueryString $request.Url.Query
                $orderIdsRaw = [string]$query["orderIds"]
                $singleOrderId = [string]$query["orderId"]
                $orderIds = @()

                if (-not [string]::IsNullOrWhiteSpace($orderIdsRaw)) {
                    $orderIds = $orderIdsRaw.Split(",", [System.StringSplitOptions]::RemoveEmptyEntries) |
                        ForEach-Object { $_.Trim() } |
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                        Select-Object -First 25
                } elseif (-not [string]::IsNullOrWhiteSpace($singleOrderId)) {
                    $orderIds = @($singleOrderId.Trim())
                }

                if ($orderIds.Count -eq 0) {
                    Write-JsonResponse -Context $context -StatusCode 400 -Payload @{
                        ok    = $false
                        error = "Missing orderId or orderIds query parameter."
                    }
                    continue
                }

                $items = @(
                    $orderIds | ForEach-Object {
                        $status = Get-OrderStatus -OrderId $_
                        [pscustomobject]@{
                            orderId   = $_
                            status    = $status.code
                            statusLabel = $status.label
                            updatedAt = $status.updatedAt
                            updatedBy = $status.updatedBy
                        }
                    }
                )

                Write-JsonResponse -Context $context -StatusCode 200 -Payload @{
                    ok     = $true
                    items  = $items
                }
                continue
            }

            if ($request.HttpMethod -ne "POST" -or $path -ne "/api/orders") {
                Write-JsonResponse -Context $context -StatusCode 404 -Payload @{
                    ok    = $false
                    error = "Route not found."
                }
                continue
            }

            $rawBody = Read-RequestBody -Request $request
            $order = ConvertTo-ValidatedOrder -RawBody $rawBody
            $orderId = New-OrderId

            $backup = [pscustomobject]@{
                orderId    = $orderId
                receivedAt = (Get-Date).ToString("o")
                payload    = $order
            }

            Save-OrderBackup -OrderId $orderId -Order $backup
            Set-OrderStatus -OrderId $orderId -StatusCode "payment_pending"
            $message = Build-OrderMessage -OrderId $orderId -Order $order
            Send-TelegramOrder -Message $message -OrderId $orderId
            $orderStatus = Get-OrderStatus -OrderId $orderId

            Write-JsonResponse -Context $context -StatusCode 200 -Payload @{
                ok          = $true
                orderId     = $orderId
                sentAt      = (Get-Date).ToString("o")
                status      = $orderStatus.code
                statusLabel = $orderStatus.label
            }
        } catch {
            Write-Warning $_.Exception.Message

            try {
                if ($context.Response -and $context.Response.OutputStream -and $context.Response.OutputStream.CanWrite) {
                    Write-JsonResponse -Context $context -StatusCode 500 -Payload @{
                        ok    = $false
                        error = $_.Exception.Message
                    }
                }
            } catch {
                # Client closed connection; continue serving next requests.
            }
        }
    }
} finally {
    $listener.Stop()
    $listener.Close()
}
