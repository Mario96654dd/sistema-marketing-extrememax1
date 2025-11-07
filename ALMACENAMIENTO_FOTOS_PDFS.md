# 📸 Almacenamiento de Fotos y PDFs en Streamlit Cloud

## ⚠️ Problema

En **Streamlit Cloud**, el sistema de archivos es **efímero** (temporal). Esto significa que:
- ❌ Los archivos guardados se **pierden** cuando la aplicación se reinicia
- ❌ Las fotos y PDFs **no se mantienen** entre sesiones
- ❌ Solo funcionan mientras la aplicación está activa

## ✅ Soluciones Disponibles

### Opción 1: GitHub como Almacenamiento (Recomendado para empezar)

**Ventajas:**
- ✅ Gratis
- ✅ Fácil de implementar
- ✅ Persistente
- ✅ Versionado automático

**Desventajas:**
- ⚠️ Límite de 100MB por archivo
- ⚠️ Límite de 1GB por repositorio (gratis)
- ⚠️ Los archivos son públicos si el repo es público

**Cómo funciona:**
- Las fotos/PDFs se guardan en el repositorio de GitHub
- Se hace commit y push automático
- Los archivos persisten entre reinicios

### Opción 2: Servicios de Almacenamiento en la Nube

#### A) Amazon S3
- ✅ Escalable
- ✅ Confiable
- ⚠️ Requiere cuenta AWS
- ⚠️ Costos según uso

#### B) Google Cloud Storage
- ✅ Integración fácil
- ✅ Generoso plan gratuito
- ⚠️ Requiere cuenta Google Cloud

#### C) Cloudinary (Para fotos)
- ✅ Gratis hasta cierto límite
- ✅ Optimización automática
- ✅ CDN incluido
- ⚠️ Solo para imágenes

### Opción 3: Base64 en Excel (No recomendado)

- ⚠️ Archivos Excel muy grandes
- ⚠️ Lento
- ⚠️ Solo para archivos pequeños

## 🚀 Implementación Recomendada: GitHub

### Configuración Necesaria

1. **Instalar GitPython:**
   ```bash
   pip install gitpython
   ```

2. **Configurar GitHub Token:**
   - Crear un Personal Access Token en GitHub
   - Guardarlo en Streamlit Secrets

3. **Modificar el código** para guardar en GitHub automáticamente

### Estructura de Carpetas en GitHub

```
tu-repositorio/
├── sistema_marketing.py
├── REGISTRO_MARKETING.xlsx
├── fotos/
│   ├── perchas/
│   ├── comerciales/
│   └── letreros/
└── documentos/
    ├── eventos/
    └── autorizaciones/
```

## 📝 Configuración en Streamlit Cloud

### 1. Crear GitHub Token

1. Ve a GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic)
2. Genera nuevo token con permisos:
   - `repo` (acceso completo a repositorios)
3. Copia el token

### 2. Configurar Streamlit Secrets

En Streamlit Cloud:
1. Ve a tu app → Settings → Secrets
2. Agrega:

```toml
[github]
token = "tu_token_aqui"
repo = "TU_USUARIO/TU_REPOSITORIO"
branch = "main"
```

## 🔧 Modificaciones al Código

El código necesita modificarse para:
1. Detectar si está en Streamlit Cloud
2. Guardar archivos en GitHub en lugar de sistema de archivos local
3. Hacer commit automático después de guardar

## ⚡ Alternativa Rápida: Usar Solo URLs

Si no quieres modificar mucho el código:
- Guardar fotos en un servicio externo (Imgur, Cloudinary)
- Guardar solo las URLs en el Excel
- Más simple pero requiere servicio externo

## 📊 Comparación de Opciones

| Opción | Costo | Complejidad | Persistencia | Recomendado |
|--------|-------|-------------|--------------|-------------|
| GitHub | Gratis | Media | ✅ Alta | ⭐⭐⭐⭐⭐ |
| S3 | Variable | Alta | ✅ Alta | ⭐⭐⭐⭐ |
| Cloudinary | Gratis* | Baja | ✅ Alta | ⭐⭐⭐⭐ |
| Base64 Excel | Gratis | Baja | ✅ Alta | ⭐⭐ |

## 🎯 Recomendación Final

**Para empezar:** Usa GitHub como almacenamiento
- Es gratis
- Funciona bien para archivos pequeños/medianos
- Fácil de implementar

**Para producción:** Considera S3 o Google Cloud Storage
- Más escalable
- Mejor rendimiento
- Más control

---

**Nota:** El código actual guarda en sistema de archivos local. Para usar en Streamlit Cloud, necesitas modificar las funciones de guardado para usar GitHub o un servicio de almacenamiento en la nube.

