param(
    [string]$BotToken = $env:TELEGRAM_BOT_TOKEN,
    [string]$AdminIds = $env:TELEGRAM_ADMIN_IDS,
    [int]$PollTimeoutSeconds = 25
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "scripts\ProductStore.ps1")

if ([string]::IsNullOrWhiteSpace($BotToken)) {
    throw "Set TELEGRAM_BOT_TOKEN before starting the bot."
}

$allowedAdmins = @()
if (-not [string]::IsNullOrWhiteSpace($AdminIds)) {
    $allowedAdmins = $AdminIds.Split(",", [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }
}

$apiBase = "https://api.telegram.org/bot$BotToken"
$offset = 0

function Send-TelegramMessage {
    param(
        [Parameter(Mandatory = $true)][string]$ChatId,
        [Parameter(Mandatory = $true)][string]$Text
    )

    $payload = @{
        chat_id = $ChatId
        text = $Text
    }

    Invoke-RestMethod -Uri "$apiBase/sendMessage" -Method Post -Body $payload | Out-Null
}

function Test-IsAllowedUser {
    param([string]$UserId)

    if ($allowedAdmins.Count -eq 0) {
        return $true
    }

    return $allowedAdmins -contains $UserId
}

function Get-HelpText {
@"
Bot commands:
/start - start and short help
/help - full command list
/categories - available categories
/list - show recent products
/add <category> | <name> | <description> | <price> | [icon]
/delete <product-id>
/rebuild - rebuild products.generated.js from JSON

Example:
/add bike | Trek Marlin 7 | Alloy frame, 29 inch wheels, 18 speeds | 34990 | fa-bicycle

Categories:
bike, moto, agro, electro
"@
}

function Get-CategoriesText {
    $defaults = Get-CategoryDefaults
    $lines = $defaults.Keys | Sort-Object | ForEach-Object {
        "$_ - $($defaults[$_].Label) ($($defaults[$_].Icon))"
    }

    return "Available categories:`n" + ($lines -join "`n")
}

function Get-ListText {
    $summary = @(Get-ProductsSummary)
    if ($summary.Count -eq 0) {
        return "Catalog is empty."
    }

    $recent = $summary | Select-Object -Last 12
    $lines = $recent | ForEach-Object {
        "$($_.id) | $($_.label) | $($_.name) | $($_.price)"
    }

    return "Recent products:`n" + ($lines -join "`n")
}

function Parse-AddCommand {
    param([string]$Text)

    $payload = $Text.Substring(4).Trim()
    if ([string]::IsNullOrWhiteSpace($payload)) {
        throw "Use format: /add category | name | description | price | [icon]"
    }

    $parts = $payload.Split("|") | ForEach-Object { $_.Trim() }
    if ($parts.Count -lt 4) {
        throw "Not enough fields. Format: category | name | description | price | [icon]"
    }

    $category = $parts[0].ToLowerInvariant()
    $name = $parts[1]
    $description = $parts[2]
    $priceRaw = ($parts[3] -replace "[^\d,\.]", "").Replace(",", ".")
    $icon = if ($parts.Count -ge 5) { $parts[4] } else { "" }

    if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($description)) {
        throw "Name and description cannot be empty."
    }

    [decimal]$price = 0
    if (-not [decimal]::TryParse($priceRaw, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$price)) {
        throw "Price is invalid. Example: 34990"
    }

    if ($price -le 0) {
        throw "Price must be greater than 0."
    }

    return @{
        Category = $category
        Name = $name
        Description = $description
        Price = $price
        Icon = $icon
    }
}

function Handle-Command {
    param(
        [string]$ChatId,
        [string]$UserId,
        [string]$Text
    )

    if (-not (Test-IsAllowedUser -UserId $UserId)) {
        Send-TelegramMessage -ChatId $ChatId -Text "Access denied."
        return
    }

    if ($Text -match "^/start") {
        Send-TelegramMessage -ChatId $ChatId -Text ("Product bot is online.`n`n" + (Get-HelpText))
        return
    }

    if ($Text -match "^/help") {
        Send-TelegramMessage -ChatId $ChatId -Text (Get-HelpText)
        return
    }

    if ($Text -match "^/categories") {
        Send-TelegramMessage -ChatId $ChatId -Text (Get-CategoriesText)
        return
    }

    if ($Text -match "^/list") {
        Send-TelegramMessage -ChatId $ChatId -Text (Get-ListText)
        return
    }

    if ($Text -match "^/rebuild") {
        Export-ProductsJs
        Send-TelegramMessage -ChatId $ChatId -Text "Catalog rebuilt. products.generated.js updated."
        return
    }

    if ($Text -match "^/delete\s+(.+)$") {
        $productId = $Matches[1].Trim()
        $deleted = Remove-ProductRecord -ProductId $productId
        if ($deleted) {
            Send-TelegramMessage -ChatId $ChatId -Text "Product $productId removed."
        } else {
            Send-TelegramMessage -ChatId $ChatId -Text "Product $productId not found."
        }
        return
    }

    if ($Text -match "^/add\b") {
        try {
            $parsed = Parse-AddCommand -Text $Text
            $product = New-ProductRecord `
                -Category $parsed.Category `
                -Name $parsed.Name `
                -Description $parsed.Description `
                -Price $parsed.Price `
                -Icon $parsed.Icon

            $created = Add-ProductRecord -Product $product
            $message = @(
                "Product added."
                "ID: $($created.id)"
                "Category: $($created.cat)"
                "Name: $($created.name)"
                "Price: $($created.price)"
            ) -join "`n"
            Send-TelegramMessage -ChatId $ChatId -Text $message
        } catch {
            Send-TelegramMessage -ChatId $ChatId -Text ("Could not add product.`n" + $_.Exception.Message)
        }
        return
    }

    Send-TelegramMessage -ChatId $ChatId -Text "Unknown command. Send /help."
}

Write-Host "Telegram product bot is running..."
Write-Host "Catalog JSON: $((Get-ProductStorePaths).Json)"
Write-Host "Catalog JS:   $((Get-ProductStorePaths).Js)"

while ($true) {
    try {
        $updates = Invoke-RestMethod -Uri "$apiBase/getUpdates?timeout=$PollTimeoutSeconds&offset=$offset" -Method Get
        if (-not $updates.ok) {
            Start-Sleep -Seconds 2
            continue
        }

        foreach ($update in $updates.result) {
            $offset = [Math]::Max($offset, [int64]$update.update_id + 1)
            if (-not $update.message) {
                continue
            }

            $message = $update.message
            if (-not $message.text) {
                continue
            }

            $chatId = [string]$message.chat.id
            $userId = [string]$message.from.id
            Handle-Command -ChatId $chatId -UserId $userId -Text $message.text.Trim()
        }
    } catch {
        Write-Warning $_.Exception.Message
        Start-Sleep -Seconds 3
    }
}
