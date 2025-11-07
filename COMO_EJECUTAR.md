# Cómo Ejecutar el Sistema

## Método 1: Usar el archivo INICIAR.bat (MÁS FÁCIL)

1. Haz doble clic en el archivo **INICIAR.bat**
2. Se abrirá una ventana con el servidor
3. **NO CIERRES** esa ventana
4. Abre `sistema.html` en tu navegador

---

## Método 2: Desde PowerShell o CMD

### Paso 1: Abrir PowerShell o CMD
- Presiona `Windows + R`
- Escribe: `powershell` o `cmd`
- Presiona Enter

### Paso 2: Ir al directorio del proyecto
```bash
cd "C:\Users\Usuario\OneDrive - Extrememax\DOCUMENTOS\MANEJOS SISTEMA MARKETING EXTREMEMAX"
```

### Paso 3: Ejecutar el servidor
```bash
python servidor.py
```

### Paso 4: Abrir el navegador
- Ve a esa carpeta en el Explorador
- Haz doble clic en `sistema.html`

---

## Método 3: Desde Git Bash (lo que estás usando)

### Paso 1: Ir al directorio correcto
```bash
cd "/c/Users/Usuario/OneDrive - Extrememax/DOCUMENTOS/MANEJOS SISTEMA MARKETING EXTREMEMAX"
```

### Paso 2: Verificar que estás en el directorio correcto
```bash
ls
```
Deberías ver archivos como: `servidor.py`, `sistema.html`, `INICIAR.bat`

### Paso 3: Ejecutar el servidor
```bash
python servidor.py
```

---

## Verificar que Funciona

Si ves esto en la terminal, el servidor está corriendo:
```
============================================
SERVIDOR MARKETING EXTREMEMAX
============================================
✅ Excel creado: REGISTRO_MARKETING.xlsx
🌐 Servidor: http://localhost:5000
============================================
 * Running on http://localhost:5000
```

---

## IMPORTANTE

- **NO CIERRES** la ventana del servidor mientras uses el sistema
- Si cierras la ventana, el servidor se detiene y no podrás guardar datos
- Para detenerlo, presiona `Ctrl + C` en la ventana del servidor

