$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$bridgeHost = "127.0.0.1"
$bridgePort = 8765
$bridgeProc = $null
$exitCode = 0
$currentStep = "init"

function Test-PortOpen {
  param(
    [string]$HostName,
    [int]$Port
  )
  try {
    $client = New-Object System.Net.Sockets.TcpClient
    $iar = $client.BeginConnect($HostName, $Port, $null, $null)
    $ok = $iar.AsyncWaitHandle.WaitOne(300)
    if (-not $ok) {
      $client.Close()
      return $false
    }
    $client.EndConnect($iar) | Out-Null
    $client.Close()
    return $true
  } catch {
    return $false
  }
}

function Stop-ManagedProcess {
  param([System.Diagnostics.Process]$Proc)
  if ($null -eq $Proc) { return }
  try {
    if (-not $Proc.HasExited) {
      Stop-Process -Id $Proc.Id -Force
    }
  } catch {}
}

function Get-PythonCommand {
  $python = Get-Command python -ErrorAction SilentlyContinue
  if ($python -and (Test-Path $python.Source)) {
    return @{ Exe = $python.Source; Args = @() }
  }
  $py = Get-Command py -ErrorAction SilentlyContinue
  if ($py -and (Test-Path $py.Source)) {
    return @{ Exe = $py.Source; Args = @("-3") }
  }
  throw "python/py not found. Please install Python and add it to PATH."
}

function Open-PageInDefaultBrowser {
  param([string]$IndexPath)
  $fullPath = (Resolve-Path $IndexPath).Path
  $uri = [System.Uri]::new($fullPath).AbsoluteUri

  $pf86 = $env:ProgramFiles
  if (Test-Path Env:'ProgramFiles(x86)') {
    $pf86 = (Get-Item Env:'ProgramFiles(x86)').Value
  }
  $browserCandidates = @(
    (Join-Path $pf86 "Microsoft\Edge\Application\msedge.exe"),
    (Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe"),
    (Join-Path $env:LocalAppData "Microsoft\Edge\Application\msedge.exe"),
    (Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe"),
    (Join-Path $pf86 "Google\Chrome\Application\chrome.exe"),
    (Join-Path $env:LocalAppData "Google\Chrome\Application\chrome.exe")
  )

  foreach ($browser in $browserCandidates) {
    if (Test-Path $browser) {
      try {
        Start-Process -FilePath $browser -ArgumentList $uri | Out-Null
        return $uri
      } catch {}
    }
  }

  try {
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c","start","", "`"$fullPath`"" -WindowStyle Hidden | Out-Null
    return $uri
  } catch {}

  try {
    Start-Process -FilePath $fullPath | Out-Null
    return $uri
  } catch {}

  throw "Unable to open browser automatically. Please open index.html manually."
}

try {
  $currentStep = "resolve_python"
  $py = Get-PythonCommand

  $bridgeScript = Join-Path $baseDir "src\bridge_server.py"
  if (-not (Test-Path $bridgeScript)) {
    throw "Bridge script not found: $bridgeScript"
  }

  $pyArgs = @()
  $pyArgs += $py.Args
  $pyArgs += $bridgeScript

  $currentStep = "start_bridge"
  $bridgeProc = Start-Process -FilePath $py.Exe -ArgumentList $pyArgs -WorkingDirectory $baseDir -PassThru -WindowStyle Hidden

  $currentStep = "wait_bridge"
  $ready = $false
  for ($i = 0; $i -lt 60; $i++) {
    if (Test-PortOpen -HostName $bridgeHost -Port $bridgePort) {
      $ready = $true
      break
    }
    Start-Sleep -Milliseconds 200
  }
  if (-not $ready) {
    throw "Bridge server failed to start on port $bridgePort."
  }

  $currentStep = "open_page"
  $indexPath = Join-Path $baseDir "index.html"
  if (-not (Test-Path $indexPath)) {
    throw "index.html not found: $indexPath"
  }
  $indexUri = Open-PageInDefaultBrowser -IndexPath $indexPath

  $currentStep = "running"
  Write-Host "Opened in default browser: $indexUri" -ForegroundColor Green
  Write-Host "Bridge server running: http://$bridgeHost`:$bridgePort"
  Write-Host "Press Enter here to stop the bridge server after you finish."
  [void][System.Console]::ReadLine()
} catch {
  $exitCode = 1
  Write-Host "Startup failed at step [$currentStep]: $($_.Exception.Message)" -ForegroundColor Red
  Write-Host "Press any key to exit..."
  try {
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
  } catch {
    Start-Sleep -Seconds 8
  }
} finally {
  Stop-ManagedProcess -Proc $bridgeProc
  exit $exitCode
}
