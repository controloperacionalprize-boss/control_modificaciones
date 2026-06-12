# Auditoría de SharePoint — Documentación

## ¿Qué hace?

Aplicación Streamlit que lista todos los archivos y carpetas de una o varias bibliotecas de SharePoint, registrando quién creó/modificó cada ítem y cuándo. Usa la **API Delta de Microsoft Graph** para sincronización incremental: la primera vez trae todo; en corridas posteriores solo descarga lo que cambió.

---

## Arquitectura general

```
Usuario (navegador)
      │
      ▼
 Streamlit UI
      │
      ├── MSAL (device-flow OAuth2) ──► Microsoft 365 / Azure AD
      │
      ├── Microsoft Graph API
      │       ├── GET /sites/{hostname}:{path}         → obtiene site_id
      │       ├── GET /sites/{site_id}/drive            → obtiene drive_id
      │       ├── GET /drives/{drive_id}/root:/{path}   → obtiene item_id
      │       └── GET /drives/{drive_id}/items/{id}/delta → items + deltaLink
      │
      └── Cache local (carpeta cache_auditoria/)
              ├── {etiqueta}_data.csv    → estado persistido del inventario
              └── {etiqueta}_delta.json  → deltaLink para la próxima sincronización
```

---

## Flujo de ejecución

### 1. Autenticación (device flow)
- Se usa `msal.PublicClientApplication` con el **Microsoft Azure CLI** (app pública registrada por Microsoft, no requiere secreto propio).
- El usuario abre una URL en el navegador e ingresa un código de un solo uso.
- El token de acceso se guarda en `st.session_state.access_token` (vive solo en la sesión activa).

### 2. Configuración de fuentes
El usuario define una tabla con columnas:

| Campo | Descripción |
|-------|-------------|
| `etiqueta` | Nombre descriptivo del origen (sirve como clave de cache) |
| `site_path` | Ruta del sitio SharePoint, ej: `/sites/OficinasPrizePeru` |
| `ruta_carpeta` | Carpeta dentro de la biblioteca de documentos, ej: `Recepción de Documentos` |

### 3. Sincronización con Delta API

#### Primera ejecución (full sync)
```
GET /drives/{drive_id}/items/{item_id}/delta
→ devuelve TODOS los ítems + @odata.deltaLink
```
El `deltaLink` se guarda en `{etiqueta}_delta.json`.

#### Ejecuciones posteriores (incremental)
```
GET {deltaLink guardado}
→ devuelve solo ítems nuevos, modificados o eliminados + nuevo deltaLink
```

#### Paginación
Si la respuesta incluye `@odata.nextLink`, el código la sigue automáticamente hasta agotar todas las páginas.

### 4. Procesamiento de cambios (`aplicar_cambios`)

Para cada ítem devuelto por la API:
- Si tiene la propiedad `"deleted"` → se marca para eliminar del DataFrame local
- Si no → se convierte en fila con los campos de auditoría

Luego:
1. Se quitan del DataFrame existente los ítems eliminados
2. Se quitan los ítems que van a ser reemplazados (actualizados)
3. Se concatenan las filas nuevas/actualizadas

### 5. Cache local
- **CSV** (`{etiqueta}_data.csv`): snapshot completo del inventario tras cada sincronización
- **JSON** (`{etiqueta}_delta.json`): contiene el `deltaLink` para la próxima llamada
- Botón **"Reporte completo"** llama a `limpiar_cache()` y borra ambos archivos, forzando un full sync

---

## Estructura de datos del DataFrame

| Columna | Origen en Graph API |
|---------|----------------------|
| `ItemId` | `item["id"]` |
| `Origen` | etiqueta configurada por el usuario |
| `Ruta` | `parentReference.path` + `name` |
| `Nombre` | `item["name"]` |
| `Tipo` | `"Carpeta"` si `"folder" in item`, sino `"Archivo"` |
| `Creado_Por` | `createdBy.user.displayName` |
| `Email_Creador` | `createdBy.user.email` |
| `Fecha_Creacion` | `createdDateTime` |
| `Modificado_Por` | `lastModifiedBy.user.displayName` |
| `Email_Modifico` | `lastModifiedBy.user.email` |
| `Fecha_Modificacion` | `lastModifiedDateTime` |
| `Tamano_Bytes` | `item["size"]` |

---

## Módulos y funciones

### Cache
| Función | Descripción |
|---------|-------------|
| `_slug(texto)` | Normaliza texto a nombre de archivo seguro (solo `a-z`, `0-9`, `_`, `-`) |
| `cache_paths(etiqueta)` | Devuelve las rutas de los dos archivos de cache para una etiqueta |
| `cargar_cache(etiqueta)` | Carga el DataFrame CSV y el deltaLink JSON; `(None, None)` si no existen |
| `guardar_cache(etiqueta, df, delta_link)` | Persiste el estado tras cada sincronización |
| `limpiar_cache(etiqueta)` | Borra los archivos de cache para forzar full sync |

### Graph API
| Función | Descripción |
|---------|-------------|
| `obtener_site_y_drive(headers, site_path)` | Resuelve `site_path` → `(site_id, drive_id)` |
| `obtener_item_id(drive_id, item_path, headers)` | Resuelve la ruta de carpeta → `item_id` |
| `obtener_delta(drive_id, item_id, headers, delta_link)` | Ejecuta la consulta delta (sigue paginación); devuelve `(items, nuevo_delta_link)` |

### Procesamiento
| Función | Descripción |
|---------|-------------|
| `item_a_fila(item, etiqueta, ruta_base)` | Mapea un ítem de Graph a una fila del DataFrame; `None` si es la raíz |
| `aplicar_cambios(df_actual, items, etiqueta, ruta_base)` | Aplica upserts y deletes sobre el DataFrame existente |
| `generar_excel(df)` | Exporta el DataFrame a un BytesIO con formato: cabecera azul, filtros, columnas autoajustadas |

---

## Instalación y uso

```bash
pip install streamlit msal requests pandas openpyxl

streamlit run auditoria_sharepoint.py
```

---

## Configuración

Las constantes en la sección `CONFIGURACION` del script:

| Constante | Valor actual | Descripción |
|-----------|--------------|-------------|
| `CLIENT_ID` | `04b07795-...` | App pública Microsoft Azure CLI |
| `TENANT_ID` | `aquanqape.onmicrosoft.com` | Tenant de la organización |
| `SP_HOSTNAME` | `aquanqape.sharepoint.com` | Hostname del SharePoint |
| `GRAPH_URL` | `https://graph.microsoft.com/v1.0` | Base URL de la API |
| `CACHE_DIR` | `cache_auditoria` | Carpeta local donde se guardan los archivos de cache |

> **Nota:** `CLIENT_ID` es la app pública de Microsoft Azure CLI. No requiere registro en Azure AD propio, pero el tenant debe permitir este client ID. Si se requiere más control, registrar una app propia en Azure AD y reemplazar `CLIENT_ID`.

---

## Consideraciones de seguridad

- El token de acceso vive únicamente en `st.session_state` (memoria de la sesión Streamlit) y no se persiste en disco.
- Los archivos de cache contienen metadata de SharePoint pero no el contenido de los archivos.
- El device flow es adecuado para uso interactivo; para ejecución desatendida se requeriría un flujo con credenciales de aplicación (client credentials).
