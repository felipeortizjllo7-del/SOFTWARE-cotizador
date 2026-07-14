# Cotizador INNOBA Colombia DMC

Software de cotizacion de paquetes turisticos (hoteles, transportes y actividades)
para INNOBA Colombia DMC. Genera cotizaciones en PDF y las envia por correo.

## Version actual

La version instalada se muestra en el encabezado de la aplicacion (ej. `v1.1.0`).
La ultima version publicada esta en [`version.json`](version.json).

## Instalar / Actualizar

Descarga el instalador mas reciente desde
**[Releases](https://github.com/felipeortizjllo7-del/SOFTWARE-cotizador/releases/latest)**
y ejecutalo. Se instala por usuario (no pide permisos de administrador).

La aplicacion **avisa sola** cuando hay una version nueva: al abrirla compara la version
instalada contra `version.json` del repositorio y, si hay una mas reciente, ofrece
descargarla e instalarla automaticamente.

> Tus datos (clientes y configuracion de empresa) se guardan en
> `%APPDATA%\CotizadorInnoba` y **no se pierden** al actualizar.

## Publicar una nueva version (solo el administrador)

1. Configura el token de GitHub una sola vez (no se comparte, se guarda en tu equipo):
   ```
   Clic derecho en  configurar_token.ps1  ->  "Ejecutar con PowerShell"
   ```
2. Publica la nueva version:
   ```
   Doble clic en  publicar_version.bat
   ```
   Esquema de 2 digitos: 1.0 -> 1.1 -> ... -> 1.9 -> 2.0
   - `publicar_version.bat`            sube al siguiente (1.0 -> 1.1 ; 1.9 -> 2.0)
   - `publicar_version.bat --major`    salta al siguiente entero (1.3 -> 2.0)
   - `publicar_version.bat 1.5`        fija exactamente esa version
   - agrega `--notas "que cambio"` para describir la version

El script compila el `.exe`, arma el instalador, sube todo al repositorio y crea el
Release. A partir de ahi, todos los equipos con la app veran el aviso de actualizacion.

## Estructura

| Archivo | Que es |
|---|---|
| `cotizador_innoba.py` | Codigo de la aplicacion de escritorio |
| `gen_html.py` | Genera la version web (`CotizadorInnoba.html`) |
| `installer.iss` | Definicion del instalador (Inno Setup) |
| `publicar_version.py` | Automatiza compilar + instalador + publicar |
| `configurar_token.ps1` | Guarda tu token de GitHub (local, privado) |
| `version.json` | Ultima version publicada (lo lee la autoactualizacion) |
| `precios_2026.json` | Tarifas 2026 |
| `descripciones_tours.json` | Descripciones de tours |
