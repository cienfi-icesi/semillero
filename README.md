# Semillero CIENFI · Taller de papers

Aplicativo de trabajo del semillero de investigación de CIENFI (Universidad Icesi),
septiembre de 2026.

**Abrir:** https://cienfi-icesi.github.io/semillero/

## Contenido

| Archivo | Qué es |
|---|---|
|  `index.html` | El aplicativo completo. Un solo archivo, sin dependencias. Es la página que publica GitHub Pages. |
| `guia_lectura_paper.pdf` | Guía de lectura que se descarga desde la pestaña "Cómo leer un paper". Debe quedar en la raíz del repositorio, junto al HTML. |
| `avances-backend.gs` | Código de Google Apps Script que recibe los reportes y las fichas. |

## Secciones

Inicio · Literatura (47 papers con DOI verificado) · Equipos · Cómo leer un paper ·
Cronograma · Horarios · Avances.

## Avances

Los equipos registran su progreso y suben la ficha de lectura en PDF. Cada envío
se guarda en una hoja de cálculo de Google y genera una notificación por correo,
mediante el Apps Script de `avances-backend.gs`.

## Nota sobre datos

Las respuestas de la encuesta de intereses (`.xlsx`) contienen correos y teléfonos
de los estudiantes y están excluidas en `.gitignore`. No deben subirse.
