<#
.SYNOPSIS
    Pushes the Anti-Gravity GSM Automator repository to GitHub account Git-LimbadriRachrla
    (limbadri3214@gmail.com) via GitHub OAuth Device Flow + GitHub REST API v3.
    Requires NO local `git` or `python` installation.
#>

param(
    [string]$GitUserName = "Git-LimbadriRachrla",
    [string]$GitUserEmail = "limbadri3214@gmail.com",
    [string]$RepoName = "anti-gravity-gsm-automator",
    [string]$SubFolder = "",
    [switch]$Private
)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# 1. Check if GITHUB_TOKEN is already set in environment or .env
$EnvFile = Join-Path $ProjectDir ".env"
$Token = $env:GITHUB_TOKEN
if (-not $Token -and (Test-Path $EnvFile)) {
    foreach ($line in Get-Content $EnvFile) {
        if ($line -match "^\s*GITHUB_TOKEN\s*=\s*(.+)$") {
            $Token = $matches[1].Trim().Trim('"').Trim("'")
        }
    }
}

# 2. If no token is saved, initiate GitHub Browser OAuth Device Flow
if (-not $Token) {
    $ClientId = "178c6fc778ccc68e1d6a" # Official GitHub CLI OAuth App Client ID
    $DeviceHeaders = @{ "Accept" = "application/json" }
    $DeviceBody = @{
        client_id = $ClientId
        scope     = "repo"
    }

    $DeviceResp = Invoke-RestMethod -Method Post -Uri "https://github.com/login/device/code" -Headers $DeviceHeaders -Body $DeviceBody
    $UserCode = $DeviceResp.user_code
    $VerificationUri = $DeviceResp.verification_uri
    $DeviceCode = $DeviceResp.device_code
    $Interval = [int]$DeviceResp.interval
    if ($Interval -lt 5) { $Interval = 5 }

    [Console]::Out.WriteLine("================================================================")
    [Console]::Out.WriteLine(" GITHUB BROWSER AUTHORIZATION REQUIRED FOR: $GitUserName")
    [Console]::Out.WriteLine(" 1. Go to: $VerificationUri")
    [Console]::Out.WriteLine(" 2. Enter one-time code: $UserCode")
    [Console]::Out.WriteLine("================================================================")
    [Console]::Out.Flush()

    # Automatically launch the user's browser to https://github.com/login/device
    try {
        Start-Process $VerificationUri
    } catch {}

    # Poll until user clicks Authorize in browser
    $MaxAttempts = 120
    for ($i = 0; $i -lt $MaxAttempts; $i++) {
        Start-Sleep -Seconds $Interval
        $PollBody = @{
            client_id   = $ClientId
            device_code = $DeviceCode
            grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
        }
        $PollResp = Invoke-RestMethod -Method Post -Uri "https://github.com/login/oauth/access_token" -Headers $DeviceHeaders -Body $PollBody
        if ($PollResp.access_token) {
            $Token = $PollResp.access_token
            Write-Host "[+] Browser authorization completed!" -ForegroundColor Green
            break
        }
        elseif ($PollResp.error -eq "slow_down") {
            $Interval += 5
        }
        elseif ($PollResp.error -eq "authorization_pending") {
            continue
        }
        else {
            throw "GitHub OAuth Device Flow error: $($PollResp.error_description)"
        }
    }
}

if (-not $Token) {
    throw "Timed out waiting for GitHub browser authorization."
}

$Headers = @{
    "Accept"               = "application/vnd.github+json"
    "Authorization"        = "Bearer $Token"
    "User-Agent"           = "Anti-Gravity-GSM-Automator"
    "X-GitHub-Api-Version" = "2022-11-28"
}

# 3. Verify authenticated GitHub account
$User = Invoke-RestMethod -Method Get -Uri "https://api.github.com/user" -Headers $Headers
$Owner = $User.login
Write-Host "[+] Authenticated as GitHub user: $Owner ($GitUserEmail)" -ForegroundColor Green

# 4. Check or Create Repository
$RepoUri = "https://api.github.com/repos/$Owner/$RepoName"
try {
    $Repo = Invoke-RestMethod -Method Get -Uri $RepoUri -Headers $Headers
    Write-Host "[+] Found existing repository: $Owner/$RepoName" -ForegroundColor Cyan
} catch {
    Write-Host "[+] Creating repository '$Owner/$RepoName'..." -ForegroundColor Cyan
    $CreateBody = @{
        name        = $RepoName
        description = "Anti-Gravity GSM Automator - Shift-aware GSM Incident Automation, Playwright Scraper, MS Teams Adaptive Cards, and Excel Ledger"
        private     = [bool]$Private
        auto_init   = $true
    } | ConvertTo-Json
    $Repo = Invoke-RestMethod -Method Post -Uri "https://api.github.com/user/repos" -Headers $Headers -Body $CreateBody -ContentType "application/json"
    Start-Sleep -Seconds 2
}

# 5. Upload all project files via GitHub Contents API
$TrackedFiles = @(
    ".gitignore",
    ".env.example",
    "config.yaml",
    "requirements.txt",
    "README.md",
    "models.py",
    "logger_setup.py",
    "roster_engine.py",
    "gsm_worker.py",
    "teams_notifier.py",
    "tracker.py",
    "main.py",
    "generate_sample_roster.py",
    "push_to_github.ps1",
    "data/.gitkeep",
    "logs/.gitkeep",
    "tests/test_roster_engine.py"
)

$AuthorMeta = @{
    name  = $Owner
    email = $GitUserEmail
}

foreach ($RelPath in $TrackedFiles) {
    $FullPath = Join-Path $ProjectDir ($RelPath -replace "/", "\")
    if (-not (Test-Path $FullPath)) { continue }

    $RemotePath = if ($SubFolder) { "$($SubFolder.Trim('/'))/$RelPath" } else { $RelPath }
    $Bytes = [System.IO.File]::ReadAllBytes($FullPath)
    $Base64Content = [System.Convert]::ToBase64String($Bytes)
    $ContentUri = "https://api.github.com/repos/$Owner/$RepoName/contents/$RemotePath"

    $Sha = $null
    try {
        $Existing = Invoke-RestMethod -Method Get -Uri $ContentUri -Headers $Headers
        $Sha = $Existing.sha
    } catch {
        $Sha = $null
    }

    $Payload = @{
        message   = "Add $RelPath - Anti-Gravity GSM Automator"
        content   = $Base64Content
        author    = $AuthorMeta
        committer = $AuthorMeta
    }
    if ($Sha) {
        $Payload["sha"] = $Sha
    }

    $JsonBody = $Payload | ConvertTo-Json -Depth 5
    Invoke-RestMethod -Method Put -Uri $ContentUri -Headers $Headers -Body $JsonBody -ContentType "application/json" | Out-Null
    [Console]::Out.WriteLine("    [OK] Pushed $RemotePath")
    [Console]::Out.Flush()
}

[Console]::Out.WriteLine("")
[Console]::Out.WriteLine("[SUCCESS] All files pushed to: https://github.com/$Owner/$RepoName")
[Console]::Out.Flush()
