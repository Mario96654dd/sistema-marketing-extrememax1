# 📤 Guía Paso a Paso para Subir a Streamlit Cloud

## Paso 1: Crear Repositorio en GitHub

### Opción A: Usando GitHub Web (Recomendado para principiantes)

1. **Ve a GitHub:**
   - Abre https://github.com en tu navegador
   - Inicia sesión (o crea una cuenta si no tienes)

2. **Crear nuevo repositorio:**
   - Haz clic en el botón **"+"** (arriba derecha) → **"New repository"**
   - **Repository name:** `sistema-marketing-extrememax`
   - **Description:** `Sistema de gestión de marketing y seguimiento de clientes`
   - **Public** (marcado) - Necesario para Streamlit Cloud gratis
   - **NO marques** "Add a README file"
   - Haz clic en **"Create repository"**

### Opción B: Usando Git desde la Terminal

1. **Abre PowerShell o CMD** en la carpeta del proyecto:
   ```powershell
   cd "C:\Users\Usuario\OneDrive - Extrememax\DOCUMENTOS\MANEJOS SISTEMA MARKETING EXTREMEMAX final"
   ```

2. **Inicializa Git (si no está inicializado):**
   ```bash
   git init
   ```

3. **Agrega los archivos:**
   ```bash
   git add sistema_marketing.py
   git add requirements.txt
   git add README.md
   git add .gitignore
   git add REGISTRO_MARKETING.xlsx
   git add EMPRESAS.xlsx
   git add logo_extrememax.png
   ```

4. **Haz el primer commit:**
   ```bash
   git commit -m "Initial commit: Sistema Marketing Extrememax"
   ```

5. **Conecta con GitHub:**
   ```bash
   git branch -M main
   git remote add origin https://github.com/TU_USUARIO/sistema-marketing-extrememax.git
   ```
   *(Reemplaza TU_USUARIO con tu nombre de usuario de GitHub)*

6. **Sube los archivos:**
   ```bash
   git push -u origin main
   ```
   *(Te pedirá usuario y contraseña/token de GitHub)*

## Paso 2: Desplegar en Streamlit Cloud

1. **Ve a Streamlit Cloud:**
   - Abre https://share.streamlit.io/
   - Haz clic en **"Sign in"**
   - Autoriza con tu cuenta de GitHub

2. **Crear nueva aplicación:**
   - Haz clic en **"New app"**
   - **Repository:** Selecciona `TU_USUARIO/sistema-marketing-extrememax`
   - **Branch:** `main`
   - **Main file path:** `sistema_marketing.py`
   - **App name:** `sistema-marketing-extrememax` (o el que prefieras)

3. **Configuración avanzada (opcional):**
   - Haz clic en **"Advanced settings"**
   - **Python version:** 3.9 o superior
   - Puedes dejar el resto por defecto

4. **Desplegar:**
   - Haz clic en **"Deploy!"**
   - Espera 2-5 minutos mientras se instala todo
   - Verás el progreso en tiempo real

5. **¡Listo!**
   - Una vez terminado, tendrás una URL como:
   - `https://sistema-marketing-extrememax.streamlit.app`
   - Haz clic en "Open app" para ver tu aplicación

## Paso 3: Actualizar la Aplicación

Cada vez que hagas cambios:

1. **Guarda los cambios en tu código**

2. **Sube a GitHub:**
   ```bash
   git add .
   git commit -m "Descripción de los cambios"
   git push
   ```

3. **Streamlit Cloud detectará los cambios automáticamente** y volverá a desplegar

## 🔧 Solución de Problemas

### Error: "Module not found"
- Verifica que `requirements.txt` tenga todas las dependencias
- Revisa los logs en Streamlit Cloud para ver qué falta

### Error: "File not found"
- Asegúrate de que los archivos Excel estén en el repositorio
- Verifica las rutas en el código

### La aplicación no se actualiza
- Espera unos minutos (puede tardar)
- Revisa los logs en Streamlit Cloud
- Verifica que el push a GitHub fue exitoso

### Problemas con archivos grandes
- Los archivos Excel grandes pueden causar problemas
- Considera optimizar o dividir los archivos

## 📝 Notas Importantes

- ⚠️ Los archivos Excel estarán públicos en GitHub
- ⚠️ Todos los usuarios compartirán los mismos datos
- ✅ Los cambios se guardan en tiempo real
- ✅ Puedes acceder desde cualquier dispositivo con internet

## 🆘 Ayuda Adicional

- Documentación Streamlit Cloud: https://docs.streamlit.io/streamlit-community-cloud
- Foro de Streamlit: https://discuss.streamlit.io/

