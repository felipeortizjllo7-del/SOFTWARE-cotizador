# ============================================================================
#  Configurar el token de GitHub para publicar versiones automaticamente.
#  Ejecutalo UNA sola vez (o cuando cambies el token).
#
#  Como obtener el token:
#   1. Entra a https://github.com/settings/tokens?type=beta  (token fine-grained)
#      - "Resource owner": tu cuenta felipeortizjllo7-del
#      - "Repository access": Only select repositories -> SOFTWARE-cotizador
#      - Permisos (Repository permissions):
#            * Contents  -> Read and write
#            * Metadata  -> Read-only (se pone solo)
#      - Genera el token (empieza por github_pat_...)
#   2. Copia ese token y pegalo aqui abajo cuando se te pida.
#
#  El token se guarda SOLO en tu equipo (%APPDATA%\CotizadorInnoba\gh_token.txt)
#  y nunca se sube al repositorio.
# ============================================================================

$dir  = Join-Path $env:APPDATA "CotizadorInnoba"
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$file = Join-Path $dir "gh_token.txt"

Write-Host ""
Write-Host "=== Configurar token de GitHub (Cotizador INNOBA) ===" -ForegroundColor Cyan
Write-Host "Pega tu token con CLIC DERECHO o Ctrl+V y presiona Enter."
Write-Host "Por seguridad NO se vera nada mientras pegas: es normal." -ForegroundColor Yellow
$sec = Read-Host "Token" -AsSecureString

$bstr  = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
$plain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

if ([string]::IsNullOrWhiteSpace($plain)) {
    Write-Host "No escribiste ningun token. Cancelado." -ForegroundColor Yellow
    exit 1
}

# Guardar sin BOM
[System.IO.File]::WriteAllText($file, $plain.Trim(), (New-Object System.Text.UTF8Encoding($false)))

Write-Host ""
Write-Host "Token guardado en: $file" -ForegroundColor Green
Write-Host "Listo. Ahora puedes publicar una version con:  publicar_version.bat" -ForegroundColor Green
