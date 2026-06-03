#Requires -Version 5.1
<#
.SYNOPSIS
    Outlookの連絡先メールアドレスのドメインを変更するスクリプト
.DESCRIPTION
    bsk-z.or.jp ドメインのメールアドレスを dfeii.or.jp に一括変更します。
    変更前にバックアップを推奨します（Outlookの連絡先を .csv にエクスポートしてください）。
.NOTES
    実行前に Outlook を終了してください。
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$OldDomain = "bsk-z.or.jp",
    [string]$NewDomain = "dfeii.or.jp",

    # 変更を行わず確認のみする場合は -WhatIf を付けて実行
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---- ログ出力ヘルパー ----
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "INFO"    { "Cyan" }
        "CHANGE"  { "Green" }
        "SKIP"    { "Gray" }
        "WARN"    { "Yellow" }
        "ERROR"   { "Red" }
    }
    Write-Host "[$timestamp][$Level] $Message" -ForegroundColor $color
}

# ---- ドメイン変換 ----
function Convert-EmailDomain {
    param([string]$Email)
    if ($Email -match "^(.+)@$([regex]::Escape($OldDomain))$") {
        return "$($Matches[1])@$NewDomain"
    }
    return $null  # 対象外
}

# ---- Outlook 起動確認 ----
$runningOutlook = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
if ($runningOutlook) {
    Write-Log "Outlook が起動中です。スクリプトの実行前に Outlook を終了してください。" "WARN"
    $answer = Read-Host "続行しますか？ (y/N)"
    if ($answer -notmatch "^[Yy]$") {
        Write-Log "処理を中断しました。" "INFO"
        exit 1
    }
}

# ---- Outlook COM オブジェクト生成 ----
Write-Log "Outlook に接続しています..."
try {
    $outlook  = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
} catch {
    Write-Log "Outlook の COM オブジェクト生成に失敗しました。Outlook がインストールされているか確認してください。`n$_" "ERROR"
    exit 1
}

# ---- 連絡先フォルダ取得 ----
# olFolderContacts = 10
$contactsFolder = $namespace.GetDefaultFolder(10)
$allItems       = $contactsFolder.Items

Write-Log "連絡先フォルダ: $($contactsFolder.Name)  ($($allItems.Count) 件)"

# ---- 変更カウンタ ----
$changedContacts = 0
$changedAddresses = 0
$results = [System.Collections.Generic.List[PSCustomObject]]::new()

# ---- 連絡先を走査 ----
foreach ($item in $allItems) {
    # 連絡先以外（配布リストなど）はスキップ
    if ($item.Class -ne 40) { continue }  # olContact = 40

    $displayName = $item.FullName
    $modified    = $false

    foreach ($field in @("Email1Address", "Email2Address", "Email3Address")) {
        $current = $item.$field
        if ([string]::IsNullOrEmpty($current)) { continue }

        $new = Convert-EmailDomain -Email $current
        if ($null -eq $new) {
            Write-Log "  スキップ: $displayName / $field = $current" "SKIP"
            continue
        }

        $results.Add([PSCustomObject]@{
            連絡先名 = $displayName
            フィールド = $field
            変更前 = $current
            変更後 = $new
        })

        if ($DryRun) {
            Write-Log "  [DryRun] $displayName / $field : $current → $new" "CHANGE"
        } else {
            Write-Log "  変更: $displayName / $field : $current → $new" "CHANGE"
            $item.$field = $new
            $modified     = $true
            $changedAddresses++
        }
    }

    if ($modified) {
        $item.Save()
        $changedContacts++
    }
}

# ---- サマリー ----
Write-Host ""
Write-Log "===== 処理完了 =====" "INFO"

if ($DryRun) {
    Write-Log "DryRun モード: 実際の変更は行いませんでした。" "WARN"
    Write-Log "変更対象アドレス数: $($results.Count) 件" "INFO"
} else {
    Write-Log "変更した連絡先数 : $changedContacts 件" "INFO"
    Write-Log "変更したアドレス数: $changedAddresses 件" "INFO"
}

# ---- 結果一覧をCSV出力 ----
if ($results.Count -gt 0) {
    $csvPath = Join-Path $PSScriptRoot "OutlookDomainChange_result_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8BOM
    Write-Log "変更一覧を出力しました: $csvPath" "INFO"
} else {
    Write-Log "変更対象の連絡先はありませんでした。" "INFO"
}

# ---- COM オブジェクト解放 ----
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($allItems)  | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($contactsFolder) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)   | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Log "完了。" "INFO"
