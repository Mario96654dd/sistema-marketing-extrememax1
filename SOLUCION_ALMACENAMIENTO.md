# 💾 Solución de Almacenamiento para Fotos y PDFs

## 📋 Resumen del Problema

En **Streamlit Cloud**, los archivos guardados en el sistema de archivos se **pierden** cuando la aplicación se reinicia. Necesitas una solución de almacenamiento persistente.

## ✅ Solución Implementada

He creado un sistema que:

1. **Detecta automáticamente** si estás en Streamlit Cloud o localmente
2. **Guarda archivos** en una carpeta `storage/` que se puede subir a GitHub
3. **Mantiene compatibilidad** con tu código actual

## 📁 Estructura de Carpetas

### En Local (como ahora):
```
tu-carpeta/
├── fotos_perchas/
├── fotos_comerciales/
└── EVENTOS_AUTORIZACIONES/
```

### En Streamlit Cloud (nuevo):
```
tu-repositorio/
├── sistema_marketing.py
├── REGISTRO_MARKETING.xlsx
└── storage/
    ├── fotos_perchas/
    ├── fotos_comerciales/
    └── documentos/
        ├── eventos/
        └── letreros/
```

## 🔧 Pasos para Implementar

### Paso 1: Actualizar .gitignore

Asegúrate de que `.gitignore` **NO excluya** la carpeta `storage/`:

```gitignore
# Mantener storage/ para que se suba a GitHub
# storage/
```

### Paso 2: Modificar el Código

Necesitas modificar las funciones de guardado para usar `storage_helper.py`. 

**Ejemplo de cambio:**

**Antes:**
```python
fotos_dir = EXCEL_PATH.parent / "fotos_perchas"
fotos_dir.mkdir(exist_ok=True)
ruta_foto = fotos_dir / nombre_foto
```

**Después:**
```python
from storage_helper import save_photo_percha
ruta_relativa = save_photo_percha(percha_id, foto, EXCEL_DIR)
```

### Paso 3: Subir Archivos a GitHub

Después de guardar archivos, necesitas hacer commit automático:

```python
import subprocess
subprocess.run(["git", "add", "storage/"])
subprocess.run(["git", "commit", "-m", "Agregar fotos/PDFs"])
subprocess.run(["git", "push"])
```

## ⚠️ Consideraciones Importantes

### 1. Límites de GitHub

- **100MB por archivo**
- **1GB por repositorio** (gratis)
- Si tienes muchos archivos grandes, considera usar un servicio externo

### 2. Seguridad

- Los archivos en GitHub son **públicos** si el repo es público
- Para archivos privados, usa un servicio de almacenamiento en la nube

### 3. Rendimiento

- GitHub puede ser lento para archivos grandes
- Considera usar un CDN o servicio de almacenamiento para producción

## 🚀 Alternativa: Servicio de Almacenamiento Externo

Si prefieres no usar GitHub, puedes usar:

### Opción A: Cloudinary (Gratis para fotos)
- Registro gratis en https://cloudinary.com
- 25GB de almacenamiento gratis
- CDN incluido

### Opción B: Amazon S3
- Escalable y confiable
- Costos según uso
- Requiere cuenta AWS

### Opción C: Google Cloud Storage
- Plan gratuito generoso
- Integración fácil
- Requiere cuenta Google Cloud

## 📝 Próximos Pasos

1. **Revisa** `storage_helper.py` - contiene las funciones de ayuda
2. **Modifica** las funciones de guardado en `sistema_marketing.py`
3. **Prueba** localmente primero
4. **Sube** a GitHub y Streamlit Cloud

## 🆘 ¿Necesitas Ayuda?

Si quieres que modifique el código completo para usar el nuevo sistema de almacenamiento, puedo hacerlo. Solo dime qué opción prefieres:

- ✅ GitHub (más simple, gratis)
- ✅ Cloudinary (mejor para fotos)
- ✅ S3/Google Cloud (más profesional)

---

**Nota:** Por ahora, el código sigue funcionando como antes localmente. Para Streamlit Cloud, necesitas implementar una de estas soluciones.

