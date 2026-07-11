$ErrorActionPreference = 'Stop'
Set-Location "c:\Users\scjsh\OneDrive\Documents\Yosafat\processing\scj-dashboard-flow"

# Login
$login = Invoke-RestMethod -Uri http://localhost:2099/api/login -Method Post -ContentType "application/json" -Body '{"username":"adminscj01","password":"adminscj02"}'
$token = $login.token
Write-Host "Token: $($token.Substring(0,16))..."

# Read file as base64
$p = Resolve-Path "data\Bookings - By Carrier - 2026.xlsx"
$bytes = [IO.File]::ReadAllBytes($p)
Write-Host "File bytes: $($bytes.Length)"
$b64 = [Convert]::ToBase64String($bytes)
Write-Host "B64 length: $($b64.Length)"

# Build body
$body = '{"file":"data:application/octet-stream;base64,' + $b64 + '"}'
Write-Host "Body length: $($body.Length)"

# Write to file
$tempFile = [IO.Path]::GetTempFileName()
[IO.File]::WriteAllText($tempFile, $body, [Text.Encoding]::UTF8)
Write-Host "Temp file: $tempFile ($((Get-Item $tempFile).Length) bytes)"

# Upload
$result = Invoke-RestMethod -Uri http://localhost:2099/api/upload -Method Post -ContentType "application/json" -Headers @{"x-auth-token"=$token} -InFile $tempFile
Write-Host "Upload result: ok=$($result.ok) records=$($result.records) GP=$($result.totalGP)"

Remove-Item $tempFile -Force
