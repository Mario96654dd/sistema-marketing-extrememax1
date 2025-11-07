# ✅ Configuración Completa de Cloudinary

## 🎯 Credenciales Completas:

- ✅ **Cloud Name:** `domc3luxa`
- ✅ **API Key:** `795545618353512`
- ✅ **API Secret:** `CBM0E2ZA7rMjkx8kUod_u4S5lTQ`

## 📋 Pasos Finales:

### Paso 1: Configurar en Streamlit Cloud

1. **Ve a Streamlit Cloud:**
   - https://share.streamlit.io/
   - Selecciona tu aplicación (o crea una nueva)

2. **Abre Secrets:**
   - Ve a **Settings** (Configuración)
   - Haz clic en **Secrets**

3. **Copia y pega este contenido:**

```toml
[cloudinary]
cloud_name = "domc3luxa"
api_key = "795545618353512"
api_secret = "CBM0E2ZA7rMjkx8kUod_u4S5lTQ"
```

4. **Guarda:**
   - Haz clic en **"Save"**
   - Espera 1-2 minutos mientras la aplicación se reinicia

### Paso 2: Verificar que Funciona

1. **Espera a que la app se reinicie**
2. **Intenta subir una foto** en cualquier sección:
   - Fotos de perchas
   - Fotos comerciales
   - Fotos de letreros
3. **Genera un PDF** de autorización
4. **Verifica** que se guarden correctamente

### Paso 3: Verificar en Cloudinary

1. **Ve al Dashboard de Cloudinary:**
   - https://cloudinary.com/console
2. **Ve a "Media Library"**
3. **Deberías ver** las carpetas:
   - `fotos_perchas/`
   - `fotos_comerciales/`
   - `fotos_letreros/`
   - `documentos/eventos/`
   - `documentos/letreros/`

## ✅ ¡Listo!

Una vez configurado, todas las fotos y PDFs se guardarán automáticamente en Cloudinary con:
- ✅ URLs permanentes
- ✅ CDN global (acceso rápido)
- ✅ Optimización automática
- ✅ 25 GB gratis de almacenamiento

## 🔒 Seguridad

- ✅ Las credenciales están en Streamlit Secrets (seguro)
- ✅ El archivo `.gitignore` protege archivos con credenciales
- ⚠️ NO subas `STREAMLIT_SECRETS_CONFIG.toml` a GitHub

## 🆘 Si algo no funciona:

1. **Verifica** que guardaste los secrets correctamente
2. **Revisa** que no haya espacios extra
3. **Espera** 2-3 minutos después de guardar
4. **Revisa los logs** en Streamlit Cloud para ver errores
5. **Verifica** que `cloudinary>=1.36.0` esté en `requirements.txt` (ya está ✅)

## 📝 Nota Importante

- Las fotos/PDFs antiguos (guardados localmente) seguirán funcionando
- Los nuevos archivos se guardarán en Cloudinary
- Las URLs de Cloudinary se guardan en el Excel

---

**¡Todo listo para usar Cloudinary!** 🚀

