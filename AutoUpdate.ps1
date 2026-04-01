# ======================================================
# Acer Marketing Sync Tool v1.3.2 (Bandwidth Optimized)
# GitHub ID: acer-wallpaper
# ======================================================

# 1. 基礎設定
$baseUrl = "https://raw.githubusercontent.com/acer-wallpaper/Acer-Wallpaper/main/"
$scriptUrl = $baseUrl + "AutoUpdate.ps1"
$localFolder = "C:\Acer_Marketing"
$localScript = "$localFolder\AutoUpdate.ps1"
$logFile = "$localFolder\sync_log.txt"
$modeFile = "$localFolder\mode.txt"

# 2. Google Form 設定 (請根據你的表單更新 5 個 Entry ID)
$formUrl = "https://docs.google.com/forms/d/e/【你的表單ID】/formResponse"

# --- 輔助功能：寫入日誌 ---
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    if (!(Test-Path $localFolder)) { New-Item -ItemType Directory -Path $localFolder }
    Add-Content -Path $logFile -Value "[$timestamp] $Message"
    if (Test-Path $logFile) { if ((Get-Content $logFile).Count -gt 1000) { $Message | Out-File $logFile -Encoding utf8 } }
}

# --- 輔助功能：獲取 Wi-Fi SSID ---
function Get-WiFiSSID {
    try {
        $ssid = (netsh wlan show interfaces | Select-String "^\s+SSID\s+:\s+(.+)$").Matches.Groups[1].Value.Trim()
        return if ($null -eq $ssid -or $ssid -eq "") { "Ethernet/No WiFi" } else { $ssid }
    } catch { return "Unknown" }
}

# --- 輔助功能：回報機台資訊 (5 欄位) ---
function Submit-Reporting {
    param($ModelName, $CurrentMode)
    try {
        $wifi = Get-WiFiSSID
        $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $postBody = @{
            "entry.1111111" = $env:COMPUTERNAME
            "entry.2222222" = $ModelName
            "entry.3333333" = $wifi
            "entry.4444444" = $now
            "entry.5555555" = $CurrentMode
        }
        Invoke-WebRequest -Uri $formUrl -Method Post -Body $postBody -ErrorAction SilentlyContinue
    } catch { Write-Log "Report failed." }
}

Write-Log "--- Session Started v1.3.2 ---"

# --- STEP A: 腳本自我更新 (文字比對) ---
try {
    $webClient = New-Object System.Net.WebClient
    $remoteContent = $webClient.DownloadString($scriptUrl)
    if (Test-Path $localScript) { $localContent = Get-Content $localScript -Raw }
    if ($remoteContent -and ($remoteContent -ne $localContent)) {
        $remoteContent | Out-File -FilePath $localScript -Encoding utf8 -Force
        Start-Process "powershell.exe" -ArgumentList "-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File `"$localScript`"" -WindowStyle Hidden
        exit
    }
} catch { Write-Log "Update check skipped." }

# --- STEP B: 讀取模式與回報 ---
try {
    $rawModel = (Get-CimInstance Win32_ComputerSystem).Model
    $model = $rawModel.Split(' ')[-1].Trim()
    $mode = "Full"
    if (Test-Path $modeFile) { $mode = (Get-Content $modeFile -Raw).Trim() }
    Submit-Reporting -ModelName $model -CurrentMode $mode

    if ($mode -eq "ReportOnly") {
        Write-Log "ReportOnly mode active. Finished."
        exit
    }

    # --- STEP C: 優化後的桌布下載 (HEAD 檢查) ---
    $extensions = @(".png", ".jpg", ".jpeg")
    $localCachePath = "$localFolder\latest_backup.jpg"
    $factoryDefault = "$localFolder\default_backup.jpg"
    $tempPath = "$env:TEMP\acer_sync"
    $targetFound = $false

    foreach ($ext in $extensions) {
        $targetUrl = "$baseUrl$model$ext"
        try {
            # 執行 HEAD 請求檢查遠端檔案大小
            $request = [System.Net.WebRequest]::Create($targetUrl)
            $request.Method = "HEAD"
            $request.Timeout = 5000
            $response = $request.GetResponse()
            $remoteSize = $response.ContentLength
            $response.Close()

            # 檢查本地檔案大小
            $localSize = 0
            if (Test-Path $localCachePath) { $localSize = (Get-Item $localCachePath).Length }

            # 只有大小不同時才真正下載
            if ($remoteSize -ne $localSize) {
                Write-Log "New version detected ($remoteSize bytes). Downloading..."
                Invoke-WebRequest -Uri $targetUrl -OutFile "$tempPath$ext" -TimeoutSec 15
                Copy-Item -Path "$tempPath$ext" -Destination $localCachePath -Force
            } else {
                Write-Log "Wallpaper on GitHub is same as local. Skipping download."
            }
            
            $targetFound = $true
            break
        } catch { continue }
    }

    # 套用樣式
    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name WallpaperStyle -Value "2"
    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name TileWallpaper -Value "0"

    $applyPath = if (Test-Path $localCachePath) { $localCachePath } else { $factoryDefault }
    
    if ($applyPath -and (Test-Path $applyPath)) {
        $code = @"
        using System.Runtime.InteropServices;
        public class Wallpaper {
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
        }
"@
        if (-not ([System.Type]::GetType("Wallpaper"))) { Add-Type -TypeDefinition $code }
        [Wallpaper]::SystemParametersInfo(0x0014, 0, $applyPath, 0x01)
    }
} catch { Write-Log "Error: $($_.Exception.Message)" }

Write-Log "--- Session Ended ---"