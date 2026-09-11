# Avisa a Bing que las paginas son nuevas o cambiaron (IndexNow).
#
# Ejecutar DESPUES de publicar en GitHub Pages, no antes: IndexNow comprueba
# que la clave este accesible en la web.
#
# Uso, desde la carpeta del repositorio, con clic derecho >
# "Ejecutar con PowerShell", o bien en una ventana de PowerShell:
#
#     cd C:\ruta\a\base-conocimiento
#     powershell -ExecutionPolicy Bypass -File .\enviar-indexnow.ps1
#
# Respuesta esperada: HTTP 200 o 202.

$ErrorActionPreference = 'Stop'

$clave       = 'd739ebb1888b38b517416c406a7ad8ff'
$host_       = 'co-spe-ing.github.io'
$claveUrl    = "https://$host_/base-conocimiento/$clave.txt"
$endpoint    = 'https://api.indexnow.org/IndexNow'

# La lista de URLs se lee de urls.txt, que genera build.py.
$rutaUrls = Join-Path $PSScriptRoot 'urls.txt'
if (-not (Test-Path $rutaUrls)) {
    Write-Host "No se encontro urls.txt junto al script." -ForegroundColor Red
    Write-Host "Ejecuta este archivo desde la carpeta del repositorio." -ForegroundColor Red
    exit 1
}

$urls = Get-Content $rutaUrls -Encoding UTF8 |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -ne '' }

Write-Host "URLs a enviar: $($urls.Count)"

# Comprobacion previa: la clave tiene que estar publicada y accesible.
try {
    $r = Invoke-WebRequest -Uri $claveUrl -UseBasicParsing -TimeoutSec 20
    if ($r.Content.Trim() -ne $clave) {
        Write-Host "El archivo de clave existe pero su contenido no coincide." -ForegroundColor Red
        exit 1
    }
    Write-Host "Clave verificada en $claveUrl" -ForegroundColor Green
}
catch {
    Write-Host "No se pudo leer $claveUrl" -ForegroundColor Red
    Write-Host "Publica primero el sitio en GitHub Pages y vuelve a intentarlo." -ForegroundColor Yellow
    exit 1
}

$cuerpo = [ordered]@{
    host        = $host_
    key         = $clave
    keyLocation = $claveUrl
    urlList     = $urls
} | ConvertTo-Json -Depth 3

# IndexNow exige UTF-8 explicito.
$bytes = [System.Text.Encoding]::UTF8.GetBytes($cuerpo)

try {
    $resp = Invoke-WebRequest -Uri $endpoint -Method Post -Body $bytes `
        -ContentType 'application/json; charset=utf-8' `
        -UseBasicParsing -TimeoutSec 60
    Write-Host "HTTP $($resp.StatusCode) - envio aceptado." -ForegroundColor Green
    Write-Host "Bing puede tardar varios dias en rastrear las paginas."
}
catch {
    $codigo = $null
    if ($_.Exception.Response) { $codigo = [int]$_.Exception.Response.StatusCode }
    switch ($codigo) {
        403     { Write-Host "HTTP 403 - la clave no es valida o todavia no es accesible en la web." -ForegroundColor Red }
        422     { Write-Host "HTTP 422 - alguna URL no pertenece al dominio $host_." -ForegroundColor Red }
        429     { Write-Host "HTTP 429 - demasiados envios. Espera unas horas y reintenta." -ForegroundColor Yellow }
        default { Write-Host "Error en el envio: $($_.Exception.Message)" -ForegroundColor Red }
    }
    exit 1
}
