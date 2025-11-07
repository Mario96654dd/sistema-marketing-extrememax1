# 🎯 Sistema Marketing Extrememax

Sistema profesional de gestión de marketing y seguimiento de clientes desarrollado con Streamlit.

## 📋 Características

- ✅ Gestión de clientes
- ✅ Gestión de letreros
- ✅ Activaciones y eventos
- ✅ Entrega de publicidad
- ✅ Entrega de perchas/exhibidores
- ✅ Entrega a comerciales
- ✅ Inventario de productos
- ✅ Reportes generales

## 🚀 Despliegue en Streamlit Cloud

### Requisitos Previos

1. Cuenta de GitHub
2. Cuenta de Streamlit Cloud (gratis)

### Pasos para Desplegar

#### 1. Preparar el Repositorio en GitHub

1. **Crea un nuevo repositorio en GitHub:**
   - Ve a https://github.com/new
   - Nombre: `sistema-marketing-extrememax` (o el que prefieras)
   - Descripción: "Sistema de gestión de marketing y seguimiento de clientes"
   - Marca como **Público** (necesario para la versión gratuita de Streamlit Cloud)
   - NO inicialices con README (ya tenemos uno)

2. **Sube los archivos necesarios:**
   ```bash
   git init
   git add sistema_marketing.py
   git add requirements.txt
   git add README.md
   git add .gitignore
   git add REGISTRO_MARKETING.xlsx
   git add EMPRESAS.xlsx
   git add logo_extrememax.png
   git commit -m "Initial commit: Sistema Marketing Extrememax"
   git branch -M main
   git remote add origin https://github.com/TU_USUARIO/TU_REPOSITORIO.git
   git push -u origin main
   ```

   **Nota:** Reemplaza `TU_USUARIO` y `TU_REPOSITORIO` con tus datos reales.

#### 2. Desplegar en Streamlit Cloud

1. **Ve a Streamlit Cloud:**
   - Visita https://share.streamlit.io/
   - Inicia sesión con tu cuenta de GitHub

2. **Nuevo App:**
   - Haz clic en "New app"
   - Selecciona tu repositorio: `TU_USUARIO/TU_REPOSITORIO`
   - Branch: `main`
   - Main file path: `sistema_marketing.py`

3. **Configuración (opcional):**
   - App name: `sistema-marketing-extrememax` (o el que prefieras)
   - Advanced settings:
     - Python version: 3.9 o superior

4. **Deploy:**
   - Haz clic en "Deploy!"
   - Espera a que termine el despliegue (2-5 minutos)

#### 3. Acceder a tu Aplicación

Una vez desplegado, tendrás una URL como:
```
https://TU_APP.streamlit.app
```

## 📁 Estructura de Archivos Necesarios

```
tu-repositorio/
├── sistema_marketing.py      # Archivo principal
├── requirements.txt          # Dependencias Python
├── README.md                 # Este archivo
├── .gitignore               # Archivos a ignorar
├── REGISTRO_MARKETING.xlsx  # Base de datos principal
├── EMPRESAS.xlsx            # Base de datos de empresas
└── logo_extrememax.png      # Logo (opcional)
```

## ⚠️ Consideraciones Importantes

### Archivos Excel

Los archivos Excel (`REGISTRO_MARKETING.xlsx`, `EMPRESAS.xlsx`) se subirán a GitHub y estarán disponibles en la aplicación en línea. 

**IMPORTANTE:** 
- Si contienen información sensible, considera usar variables de entorno o Streamlit Secrets
- Los archivos se actualizarán en tiempo real cuando uses la aplicación
- Cada usuario de la aplicación compartirá los mismos datos

### Límites de Streamlit Cloud (Gratis)

- ✅ Aplicaciones públicas ilimitadas
- ✅ 1 GB de RAM por aplicación
- ✅ CPU compartida
- ⚠️ Los archivos grandes pueden causar problemas

### Actualizar la Aplicación

Cada vez que hagas cambios y los subas a GitHub, Streamlit Cloud los detectará automáticamente y volverá a desplegar la aplicación.

```bash
git add .
git commit -m "Descripción de los cambios"
git push
```

## 🔒 Seguridad

- Los archivos Excel estarán visibles en el repositorio público
- Considera usar Streamlit Secrets para datos sensibles
- No subas contraseñas o información confidencial en el código

## 📞 Soporte

Para problemas o preguntas, revisa la documentación de Streamlit Cloud:
https://docs.streamlit.io/streamlit-community-cloud

---

**Desarrollado por:** Mario Ponce  
**Versión:** 1.0  
**Fecha:** 2025

