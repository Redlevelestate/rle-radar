# RLE · Radar diario de empresas

Cada día laborable, GitHub Actions descarga y filtra:

- **BORME** (todas las provincias): constituciones, ampliaciones de capital, nombramientos, cambios de objeto/domicilio, socio único.
- **BOE V-B**: proyectos renovables en información pública con su promotor.
- **Plataforma de Contratación**: suelo público adjudicado y obras grandes.
- **Prensa sectorial** (RSS).

Los resultados quedan en `data/AAAA-MM-DD/`. La tarea diaria de Claude los recoge, enriquece las mejores empresas y envía el Excel a Bernar.

Lanzar a mano: pestaña **Actions → Radar diario RLE → Run workflow**.
