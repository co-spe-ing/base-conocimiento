# Base de Conocimiento Jurídica DIAN

Sitio de consulta con el **texto íntegro en HTML** de la normativa, jurisprudencia y
plantillas jurídicas que usa el agente de Copilot.

## Por qué está hecho así

El agente de Copilot **no visita la página ni abre los PDF enlazados**. Las fuentes de
tipo *sitio web público* usan *Grounding with Bing Search*: el agente consulta el índice
de Bing restringido a la URL. Por eso:

1. El contenido de cada documento tiene que estar **como texto dentro del HTML**.
2. Las páginas tienen que estar **indexadas en Bing** para que el agente vea algo.

## Cómo instalar

Copia todo el contenido de este paquete a la raíz del repositorio `base-conocimiento`,
**sin borrar** los PDF y DOCX que ya están ahí, y haz push a `main`.

## Qué hay aquí

| Archivo | Para qué sirve |
|---|---|
| `index.html` | Índice general del sitio |
| `indice-tematico.html` | Qué norma y artículo aplica según la materia |
| `<documento>.html` | Ficha de cada documento, con su índice de partes |
| `<documento>-pNN.html` | Texto íntegro, dividido por título o capítulo |
| `sitemap.xml` | Mapa del sitio para Bing Webmaster Tools |
| `robots.txt` | Permite el rastreo, apunta al sitemap |
| `urls.txt` | Las 123 URLs, una por renglón, para el envío manual |
| `d739ebb1888b38b517416c406a7ad8ff.txt` | Clave de IndexNow |
| `enviar-indexnow.ps1` | Avisa a Bing que las páginas cambiaron — **Windows** |
| `enviar-indexnow.sh` | Lo mismo, para Linux y macOS |
| `.nojekyll` | Evita que GitHub Pages procese el sitio con Jekyll |
| `build.py` | Regenera todo el sitio a partir de los PDF y DOCX |

## Avisar a Bing (IndexNow)

Desde la carpeta del repositorio, **después** de publicar en GitHub Pages.

En Windows 11, abre PowerShell en la carpeta y ejecuta:

```powershell
powershell -ExecutionPolicy Bypass -File .\enviar-indexnow.ps1
```

En Linux o macOS:

```bash
./enviar-indexnow.sh
```

El script comprueba primero que la clave esté publicada y accesible; si no lo está,
avisa y no envía nada. Respuesta esperada: `HTTP 200` o `HTTP 202`.

## Regenerar el sitio

### Windows 11

Una sola vez, en PowerShell como administrador:

```powershell
winget install Python.Python.3.12
winget install JohnMacFarlane.Pandoc
winget install oschwartz10612.Poppler
```

Cierra y vuelve a abrir PowerShell para que tome el PATH, comprueba con
`pdftotext -v` y `pandoc -v`, y luego, desde la carpeta del repositorio:

```powershell
python build.py
```

Si `pdftotext` no queda en el PATH, añade a mano la carpeta `Library\bin` de Poppler
en *Variables de entorno → Path*.

### Linux o macOS

```bash
sudo apt install poppler-utils pandoc     # una sola vez
python3 build.py
```

El script relee todos los PDF y DOCX que estén en la carpeta y reescribe las páginas, el
sitemap y `urls.txt`.

**Para agregar un documento nuevo:** déjalo en la carpeta y añade su ficha al bloque
`DOCUMENTOS` al inicio de `build.py` (archivo, slug, título, categoría, emisor y los
alias con que la gente busca la norma). Luego vuelve a ejecutar el script, publica y
envía las URLs nuevas a Bing.

## Después de publicar

1. Registrar el sitio en Bing Webmaster Tools y verificarlo.
2. Enviar `sitemap.xml`.
3. Enviar las URLs de `urls.txt` desde *Envío de URL*.
4. Ejecutar `./enviar-indexnow.sh`.
5. A los 3–7 días comprobar en bing.com: `site:co-spe-ing.github.io/base-conocimiento`

Solo cuando esa búsqueda devuelva resultados tiene sentido poner la URL como fuente de
conocimiento del agente.

## Advertencia sobre el contenido

Los textos se reproducen con fines informativos y de consulta interna. Ante cualquier
discrepancia prevalece el texto oficial publicado por la entidad emisora.
