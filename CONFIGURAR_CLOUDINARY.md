# ☁️ Configuración de Cloudinary

## 📋 Pasos para Configurar Cloudinary

### Paso 1: Crear Cuenta en Cloudinary

1. **Ve a Cloudinary:**
   - Abre https://cloudinary.com en tu navegador
   - Haz clic en **"Sign Up for Free"**

2. **Regístrate:**
   - Completa el formulario de registro
   - Verifica tu email
   - Inicia sesión

### Paso 2: Obtener Credenciales

1. **Ve al Dashboard:**
   - Una vez dentro, verás tu **Dashboard**
   - En la parte superior verás tus credenciales:
     - **Cloud Name** (ej: `dabc123`)
     - **API Key** (ej: `123456789012345`)
     - **API Secret** (ej: `abcdefghijklmnopqrstuvwxyz`)

2. **Copia las credenciales:**
   - Guárdalas en un lugar seguro
   - **⚠️ NO las compartas públicamente**

### Paso 3: Configurar en Streamlit Cloud

#### Opción A: Usando Streamlit Secrets (Recomendado)

1. **Ve a tu aplicación en Streamlit Cloud:**
   - https://share.streamlit.io/
   - Selecciona tu aplicación
   - Ve a **Settings** → **Secrets**

2. **Agrega las credenciales:**
   ```toml
   [cloudinary]
   cloud_name = "tu_cloud_name_aqui"
   api_key = "tu_api_key_aqui"
   api_secret = "tu_api_secret_aqui"
   ```

3. **Guarda los cambios:**
   - Haz clic en **"Save"**
   - La aplicación se reiniciará automáticamente

#### Opción B: Usando Variables de Entorno (Local)

Si estás probando localmente, puedes crear un archivo `.env`:

```env
CLOUDINARY_CLOUD_NAME=tu_cloud_name_aqui
CLOUDINARY_API_KEY=tu_api_key_aqui
CLOUDINARY_API_SECRET=tu_api_secret_aqui
```

**Nota:** El archivo `.env` debe estar en `.gitignore` para no subirlo a GitHub.

### Paso 4: Verificar Configuración

1. **Ejecuta tu aplicación**
2. **Intenta subir una foto**
3. **Verifica que se guarde correctamente**

Si hay errores, revisa:
- ✅ Las credenciales están correctas
- ✅ Los secrets están guardados en Streamlit Cloud
- ✅ El paquete `cloudinary` está instalado (`pip install cloudinary`)

## 📊 Plan Gratuito de Cloudinary

### Límites del Plan Gratuito:

- ✅ **25 GB de almacenamiento**
- ✅ **25 GB de ancho de banda mensual**
- ✅ **25 millones de transformaciones mensuales**
- ✅ **CDN incluido**
- ✅ **Optimización automática de imágenes**

### Características:

- ✅ **Optimización automática:** Las imágenes se optimizan automáticamente
- ✅ **CDN global:** Acceso rápido desde cualquier lugar
- ✅ **Transformaciones:** Redimensionar, recortar, aplicar filtros
- ✅ **Formatos modernos:** Conversión automática a WebP, AVIF

## 🔒 Seguridad

### ⚠️ Importante:

- **NO subas tus credenciales a GitHub**
- **Usa Streamlit Secrets** para almacenarlas de forma segura
- **No compartas** tus credenciales públicamente

### Archivos a Ignorar:

Asegúrate de que `.gitignore` incluya:
```
.env
*.env
secrets.toml
```

## 🆘 Solución de Problemas

### Error: "Cloudinary no disponible"

**Causa:** Las credenciales no están configuradas correctamente.

**Solución:**
1. Verifica que los secrets estén en Streamlit Cloud
2. Verifica que los nombres de las variables sean correctos:
   - `cloud_name`
   - `api_key`
   - `api_secret`

### Error: "Invalid API credentials"

**Causa:** Las credenciales son incorrectas.

**Solución:**
1. Verifica que copiaste correctamente las credenciales
2. Asegúrate de que no haya espacios extra
3. Vuelve a copiar desde el Dashboard de Cloudinary

### Las fotos no se suben

**Causa:** Puede ser un problema de conexión o permisos.

**Solución:**
1. Verifica tu conexión a internet
2. Revisa los logs de Streamlit Cloud
3. Verifica que el plan gratuito no haya alcanzado sus límites

## 📝 Notas Adicionales

- **Fallback automático:** Si Cloudinary no está configurado, el sistema usará almacenamiento local
- **URLs persistentes:** Las URLs de Cloudinary son permanentes y no expiran
- **Optimización:** Las imágenes se optimizan automáticamente para web

## 🔗 Enlaces Útiles

- **Dashboard de Cloudinary:** https://cloudinary.com/console
- **Documentación:** https://cloudinary.com/documentation
- **Streamlit Secrets:** https://docs.streamlit.io/streamlit-community-cloud/deploy-your-app/secrets-management

---

**¡Listo!** Una vez configurado, todas las fotos y PDFs se guardarán automáticamente en Cloudinary.

