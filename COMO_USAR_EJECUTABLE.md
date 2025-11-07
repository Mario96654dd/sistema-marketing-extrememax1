# 🚀 Sistema Marketing Extrememax - Ejecutable

## ✅ Ejecutable Creado

El ejecutable del sistema se encuentra en: **`dist/SistemaMarketingExtrememax.exe`**

## 📋 Cómo Usar el Ejecutable

### Opción 1: Ejecutar directamente
1. Navega a la carpeta `dist`
2. Haz doble clic en `SistemaMarketingExtrememax.exe`
3. El sistema se abrirá automáticamente en tu navegador

### Opción 2: Usar el archivo batch
1. Haz doble clic en `EJECUTAR_EJECUTABLE.bat`
2. Esto ejecutará el sistema

## 🔧 Funcionamiento

El ejecutable:
- ✅ Inicia automáticamente el servidor Flask (puerto 5000)
- ✅ Inicia la interfaz web de Streamlit (puerto 8501)
- ✅ Abre tu navegador web automáticamente
- ✅ No requiere instalar Python ni dependencias

## ⚙️ Características

- **Sin instalación de Python**: Incluye todo lo necesario
- **Puerto automático**: Abre en http://localhost:8501
- **Datos persistentes**: Los archivos Excel se mantienen en la carpeta del ejecutable
- **Cierre seguro**: Presiona Ctrl+C en la ventana para detener el sistema

## 📁 Estructura de Archivos Necesaria

Para que el ejecutable funcione correctamente, necesita:
- `REGISTRO_MARKETING.xlsx` - Base de datos principal
- `EMPRESAS.xlsx` - Lista de empresas
- Carpetas para archivos:
  - `fotos_perchas_entregadas/`
  - `fotos_letreros/`
  - `fotos_eventos_realizados/`
  - `documentos_autorizacion/`
  - `LETREROS_AUTORIZACIONES/`

## 🎯 Pasos de Uso

1. **Copiar el ejecutable**: Copia `SistemaMarketingExtrememax.exe` donde quieras
2. **Copiar archivos de datos**: Asegúrate de copiar los archivos Excel necesarios
3. **Ejecutar**: Haz doble clic en el ejecutable
4. **Usar**: El sistema se abrirá en tu navegador

## 💡 Notas Importantes

- **Mantén la ventana abierta**: No cierres la ventana de consola mientras uses el sistema
- **Puerto en uso**: Si el puerto 8501 está ocupado, el sistema te mostrará un mensaje
- **Antivirus**: Algunos antivirus pueden dar alertas al ejecutar, es normal con PyInstaller
- **Primera ejecución**: Puede tardar unos segundos en iniciar

## 🔄 Actualizar el Ejecutable

Para crear una nueva versión del ejecutable:
1. Abre `crear_ejecutable.bat`
2. Espera a que termine
3. Usa el nuevo ejecutable en `dist/`

## 🆘 Solución de Problemas

### El ejecutable no se abre
- Verifica que no haya otro proceso usando el puerto
- Revisa que los archivos Excel estén en la misma carpeta

### Error al ejecutar
- Asegúrate de tener permisos de administrador
- Verifica que los archivos de datos existan

### Lentitud
- Puede ser normal en el primer inicio
- Cierra otros programas que usen memoria

