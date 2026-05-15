# Deploy en Streamlit Community Cloud

Esta guía documenta los pasos para publicar RMT Suite en
[Streamlit Community Cloud](https://share.streamlit.io) (gratis) con
autenticación de usuarios y aislamiento de datos.

## Prerrequisitos

- El repo `llpereirae-svg/RMT-claude` ya está en GitHub (privado).
- Tienes los hashes de las contraseñas provisionales en `.streamlit/users.json`
  local (este archivo está gitignored y NO se sube al repo).

## Pasos

### 1. Conectar Streamlit Cloud a tu GitHub

1. Entra a https://share.streamlit.io y haz login con tu cuenta de GitHub
   `llpereirae-svg`.
2. Autoriza a Streamlit Cloud a leer **repos privados**.

### 2. Crear la app

1. Click en **New app**.
2. Selecciona:
   - **Repository**: `llpereirae-svg/RMT-claude`
   - **Branch**: `main`
   - **Main file path**: `app.py`
   - **App URL** (opcional, personalizable): `rmt-claude` →
     `https://rmt-claude.streamlit.app`
3. **NO hagas click aún en Deploy** — primero configura los Secrets.

### 3. Configurar Secrets con los usuarios

1. En la pantalla de creación de la app, expande **Advanced settings**.
2. En el campo **Secrets**, pega el contenido de tu `.streamlit/users.json`
   convertido a formato TOML. Ejemplo:

```toml
[users."0930452024001"]
salt = "AQUI_EL_SALT_HEX_DEL_USERS_JSON"
hash = "AQUI_EL_HASH_HEX_DEL_USERS_JSON"
display_name = "Contador (admin)"
is_admin = true
must_change_password = true

[users."0916635154"]
salt = "AQUI_EL_SALT_HEX_DEL_USERS_JSON"
hash = "AQUI_EL_HASH_HEX_DEL_USERS_JSON"
display_name = "Colaborador"
is_admin = false
must_change_password = true
```

Para sacar los valores de `salt` y `hash`, abre tu `.streamlit/users.json`
local — está en formato JSON pero los strings son los mismos.

### 4. Deploy

Click en **Deploy**. La primera build tarda 2-3 minutos (instala dependencias).

Cuando termine, tu link público será `https://rmt-claude.streamlit.app`.

## Limitaciones del free tier de SCC

- **Filesystem efímero**: los SQLite (`data/clientes_*/...`) se reinician
  cada vez que el contenedor duerme y se despierta. Para producción
  multiusuario, migra a Postgres (Supabase o Neon free tier).
- **Cold starts**: la app puede tardar 30-60 segundos en levantar tras
  inactividad.
- **Si cambias tu contraseña** dentro de la app, ese cambio se perderá
  cuando SCC reinicie el contenedor (porque el archivo se re-hidrata
  desde Secrets). Para que persista, actualiza el hash en Secrets.
- **Datos sensibles**: aunque tienes login, los datos contables
  procesados viven en el contenedor. Si la suite va a procesar RUCs
  reales en producción, considera un VPS propio con disco persistente.

## Actualizar usuarios después del primer login

Si un usuario cambió su contraseña localmente, su `users.json` queda
diferente al Secret de SCC. Para sincronizar:

1. En tu local, copia el JSON actualizado de `.streamlit/users.json`
2. Conviértelo a TOML y pégalo en SCC → Settings → Secrets.
3. SCC reinicia el contenedor automáticamente con los nuevos hashes.

## Rollback / republish

- Cada push a `main` redeploya automáticamente.
- Para forzar republish: SCC dashboard → **Reboot app**.
- Para detener: SCC dashboard → **Delete app** (los Secrets se pierden).
