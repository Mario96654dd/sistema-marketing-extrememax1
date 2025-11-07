# 🚀 Guía Completa: Subir Sistema a Streamlit Cloud

## 📋 Resumen de Pasos

1. ✅ Crear cuenta en GitHub
2. ✅ Crear repositorio en GitHub
3. ✅ Subir archivos a GitHub
4. ✅ Configurar Cloudinary en Streamlit Cloud
5. ✅ Desplegar en Streamlit Cloud

---

## PASO 1: Crear Cuenta en GitHub (Si no tienes)

1. **Ve a GitHub:**
   - Abre https://github.com en tu navegador

2. **Regístrate:**
   - Haz clic en **"Sign up"**
   - Completa el formulario:
     - Usuario
     - Email
     - Contraseña
   - Verifica tu email

3. **Inicia sesión** con tu cuenta nueva

---

## PASO 2: Crear Repositorio en GitHub

1. **Ve a crear nuevo repositorio:**
   - Haz clic en el botón **"+"** (arriba derecha)
   - Selecciona **"New repository"**

2. **Configura el repositorio:**
   - **Repository name:** `sistema-marketing-extrememax`
   - **Description:** `Sistema de gestión de marketing y seguimiento de clientes`
   - **Marca como Público** (necesario para Streamlit Cloud gratis)
   - **NO marques** "Add a README file"
   - **NO marques** "Add .gitignore" (ya tenemos uno)
   - **NO marques** "Choose a license"

3. **Crea el repositorio:**
   - Haz clic en **"Create repository"**

---

## PASO 3: Subir Archivos a GitHub

### Opción A: Usando GitHub Desktop (Más Fácil) ⭐

1. **Descarga GitHub Desktop:**
   - Ve a https://desktop.github.com/
   - Descarga e instala GitHub Desktop

2. **Conecta con GitHub:**
   - Abre GitHub Desktop
   - Inicia sesión con tu cuenta de GitHub

3. **Clona el repositorio:**
   - En GitHub Desktop: **File → Clone Repository**
   - Selecciona tu repositorio `sistema-marketing-extrememax`
   - Elige una carpeta local donde guardarlo
   - Haz clic en **"Clone"**

4. **Copia tus archivos:**
   - Copia estos archivos a la carpeta del repositorio:
     - `sistema_marketing.py`
     - `cloudinary_helper.py`
     - `requirements.txt`
     - `README.md`
     - `REGISTRO_MARKETING.xlsx`
     - `EMPRESAS.xlsx`
     - `logo_extrememax.png` (si existe)
     - `.gitignore`
     - Todos los archivos `.md` de documentación

5. **Haz commit y push:**
   - En GitHub Desktop verás los archivos nuevos
   - Escribe un mensaje: `"Initial commit: Sistema Marketing Extrememax"`
   - Haz clic en **"Commit to main"**
   - Haz clic en **"Push origin"**

### Opción B: Usando Git desde PowerShell/CMD

1. **Abre PowerShell** en tu carpeta del proyecto:
   ```powershell
   cd "C:\Users\Usuario\OneDrive - Extrememax\DOCUMENTOS\MANEJOS SISTEMA MARKETING EXTREMEMAX final"
   ```

2. **Inicializa Git** (si es la primera vez):
   ```powershell
   git init
   ```

3. **Agrega los archivos necesarios:**
   ```powershell
   git add sistema_marketing.py
   git add cloudinary_helper.py
   git add requirements.txt
   git add README.md
   git add .gitignore
   git add REGISTRO_MARKETING.xlsx
   git add EMPRESAS.xlsx
   git add logo_extrememax.png
   git add *.md
   ```

4. **Haz el primer commit:**
   ```powershell
   git commit -m "Initial commit: Sistema Marketing Extrememax"
   ```

5. **Conecta con GitHub:**
   ```powershell
   git branch -M main
   git remote add origin https://github.com/TU_USUARIO/sistema-marketing-extrememax.git
   ```
   *(Reemplaza `TU_USUARIO` con tu nombre de usuario de GitHub)*

6. **Sube los archivos:**
   ```powershell
   git push -u origin main
   ```
   *(Te pedirá usuario y contraseña/token de GitHub)*

---

## PASO 4: Crear Cuenta en Streamlit Cloud

1. **Ve a Streamlit Cloud:**
   - Abre https://share.streamlit.io/

2. **Inicia sesión:**
   - Haz clic en **"Sign in"**
   - Selecciona **"Continue with GitHub"**
   - Autoriza la aplicación

---

## PASO 5: Desplegar en Streamlit Cloud

1. **Crea nueva aplicación:**
   - En Streamlit Cloud, haz clic en **"New app"**

2. **Configura la aplicación:**
   - **Repository:** Selecciona `TU_USUARIO/sistema-marketing-extrememax`
   - **Branch:** `main`
   - **Main file path:** `sistema_marketing.py`
   - **App name:** `sistema-marketing-extrememax` (o el que prefieras)

3. **Despliega:**
   - Haz clic en **"Deploy!"**
   - Espera 2-5 minutos mientras se instala todo

---

## PASO 6: Configurar Cloudinary

1. **Ve a Settings:**
   - En tu aplicación desplegada, haz clic en **"Settings"** (⚙️)

2. **Abre Secrets:**
   - Haz clic en **"Secrets"**

3. **Pega esta configuración:**
   ```toml
   [cloudinary]
   cloud_name = "domc3luxa"
   api_key = "795545618353512"
   api_secret = "CBM0E2ZA7rMjkx8kUod_u4S5lTQ"
   ```

4. **Guarda:**
   - Haz clic en **"Save"**
   - La aplicación se reiniciará automáticamente

---

## PASO 7: Verificar que Funciona

1. **Abre tu aplicación:**
   - Haz clic en **"Open app"** o ve a la URL que te dieron
   - URL será algo como: `https://sistema-marketing-extrememax.streamlit.app`

2. **Prueba las funciones:**
   - Intenta subir una foto
   - Genera un PDF
   - Verifica que todo funcione

---

## 📁 Archivos que DEBES Subir

### ✅ Archivos Necesarios:
- `sistema_marketing.py` ✅
- `cloudinary_helper.py` ✅
- `requirements.txt` ✅
- `README.md` ✅
- `.gitignore` ✅
- `REGISTRO_MARKETING.xlsx` ✅
- `EMPRESAS.xlsx` ✅
- `logo_extrememax.png` (si existe) ✅

### ❌ Archivos que NO debes subir:
- `STREAMLIT_SECRETS_CONFIG.toml` ❌ (tiene credenciales)
- `CLOUDINARY_SECRETS.toml` ❌ (tiene credenciales)
- Carpetas `fotos_*/` ❌ (muy grandes)
- `EVENTOS_AUTORIZACIONES/` ❌ (muy grandes)
- `LETREROS_AUTORIZACIONES/` ❌ (muy grandes)
- Archivos `.bat` ❌
- Archivos `.exe` ❌

---

## 🆘 Solución de Problemas

### Error: "Module not found"
- **Solución:** Verifica que `requirements.txt` tenga todas las dependencias
- Revisa los logs en Streamlit Cloud

### Error: "File not found"
- **Solución:** Asegúrate de que los archivos Excel estén en el repositorio
- Verifica las rutas en el código

### La aplicación no se actualiza
- **Solución:** Espera unos minutos
- Revisa los logs en Streamlit Cloud
- Verifica que el push a GitHub fue exitoso

### Cloudinary no funciona
- **Solución:** Verifica que los secrets estén guardados correctamente
- Revisa que no haya espacios extra en las credenciales
- Espera 2-3 minutos después de guardar

---

## ✅ Checklist Final

Antes de considerar que todo está listo:

- [ ] Repositorio creado en GitHub
- [ ] Archivos subidos a GitHub
- [ ] Aplicación desplegada en Streamlit Cloud
- [ ] Cloudinary configurado en Secrets
- [ ] Aplicación funciona correctamente
- [ ] Puedes subir fotos
- [ ] Puedes generar PDFs

---

## 🎉 ¡Listo!

Una vez completados todos los pasos, tu sistema estará disponible en línea en:
```
https://TU_APP_NAME.streamlit.app
```

Puedes acceder desde cualquier dispositivo con internet.

---

## 📞 Ayuda Adicional

- **Documentación Streamlit Cloud:** https://docs.streamlit.io/streamlit-community-cloud
- **Documentación Cloudinary:** https://cloudinary.com/documentation
- **Foro de Streamlit:** https://discuss.streamlit.io/

---

**¡Éxito con tu despliegue!** 🚀

