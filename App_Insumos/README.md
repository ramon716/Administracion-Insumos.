# Administración de Insumos – Rural Bioenergía S.A.

App web (Google Apps Script) para que los capataces:

- den de **baja insumos** del stock de su estancia y almacén, enviándolas a **Estansys** (quedan PENDIENTE hasta que contabilidad las aprueba y se registran en SAP);
- registren la **recepción de compras** aprobadas y pendientes (por ítem y en parcial);
- consulten historial, estados y stock por almacén.

## Seguridad del token

**El token de Estansys no está en ningún archivo de este repositorio.**

- Se guarda solo en Apps Script → ⚙️ Configuración del proyecto → *Propiedades de la secuencia de comandos* → `ESTANSYS_TOKEN`.
- Las llamadas a Estansys las hace el servidor de Google (`UrlFetchApp`); el navegador del usuario nunca recibe el token.
- Nunca pegues el token en `Code.gs`, `Index.html`, en un commit, issue o README. Si se filtra, pedí a Estansys que lo revoque y genere otro.

## Archivos

| Archivo | Dónde va |
|---|---|
| `Code.gs` | Apps Script → Código.gs |
| `Index.html` | Apps Script → archivo HTML `Index` |

## Planilla

Hojas: `Usuarios`, `Almacenes` (estancia → almacenes), `Bajas`, `Recepciones`.
La hoja `Usuarios` tiene las claves: no la compartas ni la subas a GitHub.

Respaldo: cada noche (23 h) se guarda una copia `.xlsx` con fecha en la carpeta de Drive
"Respaldos Administración de Insumos" (sin la hoja Usuarios). Se guardan los últimos 90 días.
