# 📸 Resumen: Almacenamiento de Fotos y PDFs

## ⚠️ El Problema

En **Streamlit Cloud**, las fotos y PDFs que guardas se **pierden** cuando la aplicación se reinicia porque el sistema de archivos es temporal.

## ✅ Soluciones Disponibles

### Opción 1: GitHub (Recomendado para empezar) ⭐

**Cómo funciona:**
- Las fotos/PDFs se guardan en una carpeta `storage/` en tu repositorio
- Se hace commit automático a GitHub
- Los archivos persisten entre reinicios

**Ventajas:**
- ✅ Gratis
- ✅ Fácil de implementar
- ✅ Persistente

**Desventajas:**
- ⚠️ Límite de 100MB por archivo
- ⚠️ Límite de 1GB por repositorio
- ⚠️ Archivos públicos si el repo es público

**Archivos creados:**
- `storage_helper.py` - Funciones para guardar archivos
- `ALMACENAMIENTO_FOTOS_PDFS.md` - Documentación completa
- `SOLUCION_ALMACENAMIENTO.md` - Guía de implementación

### Opción 2: Servicios de Nube (Para producción)

**Cloudinary** (Gratis para fotos):
- 25GB gratis
- CDN incluido
- Optimización automática

**Amazon S3 / Google Cloud Storage:**
- Escalable
- Profesional
- Requiere configuración

## 🎯 Recomendación

**Para empezar:** Usa GitHub
- Ya está preparado
- Solo necesitas modificar el código para usar `storage_helper.py`

**Para producción:** Considera Cloudinary o S3
- Mejor rendimiento
- Más escalable

## 📝 Próximos Pasos

1. **Lee** `SOLUCION_ALMACENAMIENTO.md` para detalles
2. **Decide** qué opción usar
3. **Modifica** el código para usar el almacenamiento elegido

¿Quieres que modifique el código completo para usar GitHub como almacenamiento?

