# ⚡ Configuración Rápida de Cloudinary

## ✅ Credenciales que ya tienes:

- **API Key:** `795545618353512` ✅
- **API Secret:** `CBM0E2ZA7rMjkx8kUod_u4S5lTQ` ✅
- **Cloud Name:** ⚠️ **FALTA** - Necesitas obtenerlo del Dashboard

## 📋 Pasos para Configurar:

### Paso 1: Obtener Cloud Name

1. **Ve al Dashboard de Cloudinary:**
   - https://cloudinary.com/console
   - Inicia sesión con tu cuenta

2. **Encuentra tu Cloud Name:**
   - Está en la parte superior del Dashboard
   - Es algo como: `dabc123` o `mi-empresa-123`
   - **Cópialo**

### Paso 2: Configurar en Streamlit Cloud

1. **Ve a Streamlit Cloud:**
   - https://share.streamlit.io/
   - Selecciona tu aplicación
   - O crea una nueva si aún no la tienes

2. **Abre Secrets:**
   - Ve a **Settings** (Configuración)
   - Haz clic en **Secrets**

3. **Pega este contenido:**

```toml
[cloudinary]
cloud_name = "TU_CLOUD_NAME_AQUI"
api_key = "795545618353512"
api_secret = "CBM0E2ZA7rMjkx8kUod_u4S5lTQ"
```

4. **Reemplaza `TU_CLOUD_NAME_AQUI`** con tu Cloud Name real

5. **Guarda:**
   - Haz clic en **"Save"**
   - La aplicación se reiniciará automáticamente

### Paso 3: Verificar

1. **Espera a que la app se reinicie** (1-2 minutos)
2. **Intenta subir una foto** en tu aplicación
3. **Verifica que funcione**

## 🔒 Seguridad

⚠️ **IMPORTANTE:**
- ✅ Estas credenciales están ahora en Streamlit Secrets (seguro)
- ❌ NO las subas a GitHub
- ❌ NO las compartas públicamente
- ✅ El archivo `.gitignore` ya está configurado para ignorar archivos con credenciales

## 🆘 Si algo no funciona:

1. **Verifica que el Cloud Name sea correcto**
2. **Verifica que no haya espacios extra en las credenciales**
3. **Revisa los logs de Streamlit Cloud** para ver errores
4. **Asegúrate de que `cloudinary` esté en `requirements.txt`** (ya está ✅)

## ✅ Listo!

Una vez configurado, todas las fotos y PDFs se guardarán automáticamente en Cloudinary y tendrás URLs permanentes que funcionan desde cualquier lugar.

---

**Nota:** Si aún no tienes el Cloud Name, ve al Dashboard de Cloudinary y lo encontrarás en la parte superior de la página.

