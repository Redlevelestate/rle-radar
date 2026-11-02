# RLE · Radar diario de empresas — instrucciones de cada ejecución

Se ejecuta **de lunes a viernes a las 09:10** (Europe/Madrid). Antes, a las 07:15 y 07:45, el robot de GitHub Actions del repositorio
`Redlevelestate/rle-radar` descarga y filtra el BORME, el BOE V-B, la Plataforma de Contratación y la prensa, y deja los
resultados en `data/AAAA-MM-DD/` del repositorio.
Objetivo: un **Excel con las empresas nuevas del día** del público objetivo de Red Level Estate (promotoras, industrial y logística,
renovables, centros de datos, inversores/SOCIMIs/fondos/family offices, hoteles, residencias y operadores en expansión),
con web, email y decisor cuando existan, y un **email resumen** a Bernar.

## Límites (léelos primero)

- **Por ahora NO se sube nada a Zoho ni se manda nada al validador.** Bernar ha pedido solo el Excel para ver qué hace la máquina.
  En Zoho solo se **consulta** (anti-duplicados). Nunca crear, editar ni borrar registros.
- **Nunca inventes** empresas, webs, emails ni personas. Si no lo encuentras con evidencia, la celda queda vacía.
- **Nunca emails personales** ni emails sacados de políticas de privacidad, cookies, RGPD, reclamaciones o canal ético
  (`dpo@`, `privacidad@`, `lopd@`, `rgpd@`, `protecciondedatos@`, `reclamaciones@`, `etica@`, `canaldenuncias@`…).
  Bernar no los quiere: si es el único email que aparece, la empresa va sin email.
- Excluye intermediarios: consultoras inmobiliarias, brokers, agencias, APIs, headhunters (CBRE, JLL, Savills, Colliers, Cushman, Knight Frank, etc.).
- **idealista no se usa** en esta máquina (sus condiciones prohíben recopilar anuncios de forma sistemática).
- El único email que sale es el resumen a Bernar (cuenta de Gmail conectada).
- Si es **sábado o domingo**, no hagas nada.

## Configuración

| Parámetro | Valor |
|---|---|
| Carpeta principal (Excels del día) | «RLE · Radar diario de empresas» — ID `1RNgeJhabBqyb_xc10NKbWtZl0a7vJS3N` |
| Repositorio GitHub | `Redlevelestate/rle-radar` (privado): radar.py, excel.py, este archivo y `data/AAAA-MM-DD/` con lo descargado por el robot |
| Carpeta del sistema | «Sistema (no tocar)» — ID `1vfDhYRorMBlU20MGJuZP8atoxfpfrm0J` (registros `registro_*.txt`) |
| Registro de empresas vistas | el archivo más reciente de la carpeta del sistema cuyo título empiece por `registro_` (texto, `registro_AAAA-MM-DD.txt`) |
| Nombre del Excel | `Radar_RLE_AAAA-MM-DD.xlsx` (sin convertir a Google Sheets) |
| Máximo a enriquecer al día | 45 empresas (presupuesto orientativo: unas 150 búsquedas web en total) |

## Paso 0 · Preparar

1. Fecha de hoy en Europe/Madrid (`FECHA` = AAAA-MM-DD, salvo que el mensaje de la tarea indique otra). Fin de semana → termina sin hacer nada.
2. Da acceso al repositorio con la herramienta `add_repo` (owner `Redlevelestate`, repo `rle-radar`, access `read`) y clónalo:
   `git clone --depth 1 https://github.com/Redlevelestate/rle-radar /tmp/radar` (timeout amplio). Trabaja en `/tmp/radar`.
3. Busca en la carpeta del sistema de Drive los archivos con título que empiece por `registro_`; descarga el **más reciente por título** como
   `/tmp/radar/registro.txt`. Si no hay ninguno, empieza con un archivo vacío.
4. `pip install --break-system-packages openpyxl pypdf` si faltan.
5. Lleva una lista `alertas` y guárdala al final en `/tmp/radar/out/alertas.json`.

## Paso 1 · Recoger lo que ha descargado el robot

1. Si existe `/tmp/radar/data/FECHA/`, copia sus `borme.json`, `renovables.json`, `licitaciones.json` y `rss.json` a `/tmp/radar/out/`.
2. Si no existe todavía (el robot va con retraso), espera con un bucle de comprobación (`git pull` cada 5 minutos, máximo 40 minutos).
   Si sigue sin estar, continúa sin esas fuentes y alerta: «El robot de GitHub no ha dejado los datos de hoy».
3. Si `borme.json` dice «sin BORME ese día», no es error (festivo): menciónalo en el resumen.
4. Si `rss.json` marca medios con ERROR, menciónalo solo en una línea al final del email (no es crítico: las noticias se cubren también por búsqueda web).
5. No ejecutes `radar.py` para descargar: desde aquí esas webs están bloqueadas; lo hace el robot.

## Paso 2 · Noticias, inversores y empleo (IA + búsqueda web)

Crea `out/web.json` con `{"noticias": [...], "empleo": [...]}`.

**Noticias** — parte de los titulares de `rss.json` y complementa con 4–6 búsquedas web de hoy (varía las consultas; añade
el mes y el año): compra de suelo, nueva planta/nave/centro logístico, apertura de hoteles/residencias/centros de datos,
SOCIMIs y fondos que compran activos, family offices que invierten en inmobiliario, financiaciones para desarrollos,
operadores en expansión (supermercados, gasolineras, gimnasios, self-storage, clínicas…).
Solo noticias **de los últimos 3 días** y **en España**. Una fila por empresa que **compra, invierte, se expande o busca ubicación**
(no la que vende, salvo que sea un grupo que también promueve). Campos:
`empresa, sector, señal (Compra de suelo | Nueva implantación | Expansión | Inversión | Financiación | Adquisición de activo), resumen (1 frase, tus palabras), zona, fecha, url`.

**Empleo** — 3–4 búsquedas de ofertas publicadas recientemente en España de: técnico/a de suelo, gestor/a de suelo,
responsable/director/a de expansión, técnico/a de expansión o implantación, land manager, acquisitions manager,
development manager, técnico de inversiones inmobiliarias. Solo si se ve **la empresa que contrata** (si la publica un
headhunter sin nombre de cliente, descártala). Campos: `empresa, sector, puesto, ubicacion, fecha, url`.

## Paso 3 · Elegir y enriquecer candidatos

1. `python3 excel.py candidatos out registro.txt 45` → `out/candidatos.json` (ya ordenado por interés; excluye las enriquecidas otros días).
2. Para cada candidato, en orden, busca:
   - **Web oficial.** Si el nombre es una sociedad vehículo (nombres con números, «Solar …», «Fotovoltaica …», «Proyecto …», «… Desarrollos 12»),
     intenta identificar el **grupo** que hay detrás (socio único del BORME, la noticia, el anuncio del BOE, el domicilio compartido)
     y enriquece el grupo, explicándolo en `motivo` («Vehículo de Grupo X para…»). Si no se identifica con seguridad, sin web.
   - **Email corporativo** de la web (contacto, aviso legal, pie): preferencia `expansion@`/`suelo@`/`inversiones@` > `comercial@` > `info@`/`contacto@`.
     Respeta las exclusiones de «Límites». Anota dónde se encontró en `fuente_email` (p. ej. «Web · contacto»).
     Si no aparece en la primera búsqueda, segunda pasada: página de contacto, pie de la web, sala de prensa y búsqueda `"@dominio"`.
     El email siempre tiene que estar publicado en una fuente leída, nunca deducido.
   - **Decisor del BORME**: si el candidato viene del BORME, sus administradores (columna `administradores`, o los del grupo si se ha
     identificado) son el decisor de partida: ponlos en `decisor`/`cargo` aunque no aparezcan en LinkedIn.
   - **Decisor en LinkedIn**: búsqueda de LinkedIn tipo `"<empresa>" (expansión OR suelo OR adquisiciones OR "director general" OR CEO) site:linkedin.com/in`.
     Solo si el perfil es claramente de esa empresa y el cargo actual encaja. `decisor`, `cargo`, `linkedin` (URL).
   - **Teléfono** de la web si aparece.
   - **¿Ya en Zoho?**: calcula el dominio raíz de la web (sin `www.`, sin subdominios salvo ccTLD dobles tipo `.com.es`) y consulta:
     `select id from Leads where (Email like '%@D' or web like '%D%') limit 5`,
     `select id from Contacts where Email like '%@D' limit 5`, `select id from Accounts where Website like '%D%' limit 5`.
     Confirma la coincidencia exacta de dominio. `en_crm` = «Sí (Leads/Contactos/Cuentas)» o «No». Si Zoho falla, «No comprobado».
   - **motivo**: una frase sobre **por qué escribirles ahora** (la señal concreta: «Amplía capital en 2,4 M€», «Crea un vehículo para una promoción en Girona»…).
3. **Encaje con RLE**: antes de enriquecer, salta (sin gastar búsquedas) las grandes cotizadas, bancos, aseguradoras y multinacionales
   con equipos propios de suelo (Neinor, Aedas, Merlin, Bankinter, Amazon, etc.) y las sociedades claramente ajenas al sector.
   Prioriza promotoras medianas, family offices, logística, renovables, hoteles/residencias y operadores en expansión.
4. Máximo ~3–4 búsquedas por empresa; si no aparece la web en 2 búsquedas, pasa a la siguiente.
5. Guarda `out/enriquecidas.json` como objeto `{clave_del_candidato: {empresa, sector, señal, motivo, web, email, fuente_email,
   decisor, cargo, linkedin, telefono, provincia, en_crm, origen}}` (`origen`: BORME | Renovables | Licitaciones | Noticias | Empleo).
   Guarda solo las que tengan al menos web. Las que no tengan nada no se guardan (se reintentarán si vuelven a aparecer con otra señal).

## Paso 4 · Excel y registro

1. Guarda `out/alertas.json` (lista de textos).
2. `python3 excel.py construir out registro.txt FECHA` (desde `/tmp/radar`) → `out/Radar_RLE_FECHA.xlsx`, `out/registro_nuevo.txt`, `out/resumen.json`.
3. Abre el Excel con openpyxl para comprobar que tiene las 8 hojas y que «★ Para contactar» no está vacía si hubo enriquecidas.
4. Sube el Excel a la carpeta principal (`parentId 1RNgeJhabBqyb_xc10NKbWtZl0a7vJS3N`, `contentMimeType`
   `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`, `disableConversionToGoogleType = true`). Guarda su `viewUrl`.
5. Sube `registro_nuevo.txt` a la carpeta del sistema con el título `registro_FECHA.txt` (`text/plain`, `disableConversionToGoogleType = true`).
   Después, mueve a la papelera los `registro_*.txt` de **más de 7 días** (deja siempre al menos los 3 más recientes).

## Paso 5 · Email a Bernar

Gmail, a la propia cuenta conectada, en HTML sencillo y limpio (tipografía del sistema, acento color #8C1D2C):

- **Asunto**: `Radar RLE · DD/MM · <nuevas> empresas nuevas · <contactar> para contactar` (si hay alertas, empieza por `⚠`).
- **Cuerpo**, en este orden:
  1. Alertas (si hay).
  2. Cifras del día: nuevas, para contactar, con email, para carta física; y una línea por fuente (BORME: provincias y actos leídos, señales A/B/C; renovables; licitaciones; noticias; empleo).
  3. **Las 10 mejores del día** (de `resumen.json` → `top`): tabla Empresa · Por qué ahora · Email · Decisor.
  4. Botón/enlace al Excel del día (viewUrl) y a la carpeta https://drive.google.com/drive/folders/1RNgeJhabBqyb_xc10NKbWtZl0a7vJS3N
- Sin análisis largos.

## Errores

- Si **todas** las fuentes oficiales y RSS fallan por dominio no permitido, sigue con noticias y empleo por búsqueda web,
  genera el Excel igualmente y pon la alerta en el asunto.
- Si Drive falla al subir el Excel, adjúntalo al email (base64) y alerta.
- Si Gmail falla, deja el Excel en Drive; no hay más que hacer.
- Nunca borres Excels anteriores.
