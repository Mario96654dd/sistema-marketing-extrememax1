# ⚡ Comandos Rápidos para Subir a GitHub

## Si ya tienes Git instalado, usa estos comandos:

### 1. Navegar a tu carpeta:
```powershell
cd "C:\Users\Usuario\OneDrive - Extrememax\DOCUMENTOS\MANEJOS SISTEMA MARKETING EXTREMEMAX final"
```

### 2. Inicializar Git (solo la primera vez):
```powershell
git init
```

### 3. Agregar archivos:
```powershell
git add sistema_marketing.py cloudinary_helper.py requirements.txt README.md .gitignore REGISTRO_MARKETING.xlsx EMPRESAS.xlsx logo_extrememax.png *.md
```

### 4. Hacer commit:
```powershell
git commit -m "Initial commit: Sistema Marketing Extrememax"
```

### 5. Conectar con GitHub (reemplaza TU_USUARIO):
```powershell
git branch -M main
git remote add origin https://github.com/TU_USUARIO/sistema-marketing-extrememax.git
```

### 6. Subir archivos:
```powershell
git push -u origin main
```

---

## Para actualizar después de hacer cambios:

```powershell
git add .
git commit -m "Descripción de los cambios"
git push
```

---

## Si te pide autenticación:

GitHub ya no acepta contraseñas. Necesitas un **Personal Access Token**:

1. Ve a GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic)
2. Genera nuevo token con permisos `repo`
3. Úsalo como contraseña cuando Git te la pida

