# -*- coding: utf-8 -*-
# Helper para almacenamiento en Cloudinary
# Cloudinary es un servicio de almacenamiento en la nube para imágenes, videos y documentos

import os
import cloudinary
import cloudinary.uploader
import cloudinary.api
from pathlib import Path
from datetime import datetime
import streamlit as st

# Configurar Cloudinary
def init_cloudinary():
    """Inicializar Cloudinary con credenciales desde variables de entorno o Streamlit Secrets"""
    try:
        # Intentar obtener desde Streamlit Secrets primero
        try:
            cloud_name = st.secrets.get("cloudinary", {}).get("cloud_name")
            api_key = st.secrets.get("cloudinary", {}).get("api_key")
            api_secret = st.secrets.get("cloudinary", {}).get("api_secret")
        except:
            cloud_name = None
            api_key = None
            api_secret = None
        
        # Si no están en secrets, intentar desde variables de entorno
        if not cloud_name:
            cloud_name = os.getenv("CLOUDINARY_CLOUD_NAME")
        if not api_key:
            api_key = os.getenv("CLOUDINARY_API_KEY")
        if not api_secret:
            api_secret = os.getenv("CLOUDINARY_API_SECRET")
        
        if cloud_name and api_key and api_secret:
            cloudinary.config(
                cloud_name=cloud_name,
                api_key=api_key,
                api_secret=api_secret,
                secure=True
            )
            return True
        else:
            return False
    except Exception as e:
        print(f"Error inicializando Cloudinary: {e}")
        return False

def upload_photo_to_cloudinary(file_data, folder, public_id_prefix):
    """
    Subir foto a Cloudinary
    
    Args:
        file_data: Datos del archivo (UploadedFile de Streamlit o bytes)
        folder: Carpeta en Cloudinary (ej: "fotos_perchas", "fotos_comerciales")
        public_id_prefix: Prefijo para el ID público (ej: "percha_1", "comercial_10")
    
    Returns:
        dict: Información del archivo subido con 'url' y 'public_id', o None si falla
    """
    try:
        if not init_cloudinary():
            return None
        
        # Obtener datos del archivo
        if hasattr(file_data, 'getbuffer'):
            # Es un UploadedFile de Streamlit
            file_bytes = file_data.getbuffer()
            file_format = file_data.name.split('.')[-1] if '.' in file_data.name else 'jpg'
        else:
            # Es bytes o Path
            if isinstance(file_data, Path):
                file_bytes = file_data.read_bytes()
                file_format = file_data.suffix[1:] if file_data.suffix else 'jpg'
            else:
                file_bytes = file_data
                file_format = 'jpg'
        
        # Generar ID único
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        public_id = f"{folder}/{public_id_prefix}_{timestamp}"
        
        # Subir a Cloudinary
        result = cloudinary.uploader.upload(
            file_bytes,
            folder=folder,
            public_id=public_id,
            resource_type="auto",  # Detecta automáticamente si es imagen, video o documento
            overwrite=False,
            use_filename=False
        )
        
        return {
            'url': result.get('secure_url'),
            'public_id': result.get('public_id'),
            'format': result.get('format'),
            'bytes': result.get('bytes')
        }
    except Exception as e:
        print(f"Error subiendo foto a Cloudinary: {e}")
        return None

def upload_pdf_to_cloudinary(pdf_path, folder, public_id_prefix):
    """
    Subir PDF a Cloudinary
    
    Args:
        pdf_path: Ruta del archivo PDF (Path o str)
        folder: Carpeta en Cloudinary (ej: "documentos/eventos", "documentos/letreros")
        public_id_prefix: Prefijo para el ID público
    
    Returns:
        dict: Información del archivo subido con 'url' y 'public_id', o None si falla
    """
    try:
        if not init_cloudinary():
            return None
        
        pdf_path_obj = Path(pdf_path)
        if not pdf_path_obj.exists():
            return None
        
        # Leer PDF
        pdf_bytes = pdf_path_obj.read_bytes()
        
        # Generar ID único
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        public_id = f"{folder}/{public_id_prefix}_{timestamp}"
        
        # Subir a Cloudinary como documento
        result = cloudinary.uploader.upload(
            pdf_bytes,
            folder=folder,
            public_id=public_id,
            resource_type="raw",  # PDFs se suben como raw
            overwrite=False,
            use_filename=False
        )
        
        return {
            'url': result.get('secure_url'),
            'public_id': result.get('public_id'),
            'bytes': result.get('bytes')
        }
    except Exception as e:
        print(f"Error subiendo PDF a Cloudinary: {e}")
        return None

def delete_from_cloudinary(public_id, resource_type="image"):
    """
    Eliminar archivo de Cloudinary
    
    Args:
        public_id: ID público del archivo
        resource_type: Tipo de recurso ("image", "video", "raw")
    
    Returns:
        bool: True si se eliminó correctamente
    """
    try:
        if not init_cloudinary():
            return False
        
        result = cloudinary.uploader.destroy(public_id, resource_type=resource_type)
        return result.get('result') == 'ok'
    except Exception as e:
        print(f"Error eliminando archivo de Cloudinary: {e}")
        return False

def get_cloudinary_url(public_id, transformation=None):
    """
    Obtener URL de Cloudinary con transformaciones opcionales
    
    Args:
        public_id: ID público del archivo
        transformation: Diccionario con transformaciones (ej: {"width": 800, "quality": "auto"})
    
    Returns:
        str: URL del archivo
    """
    try:
        if not init_cloudinary():
            return None
        
        if transformation:
            return cloudinary.CloudinaryImage(public_id).build_url(**transformation)
        else:
            return cloudinary.CloudinaryImage(public_id).build_url()
    except Exception as e:
        print(f"Error obteniendo URL de Cloudinary: {e}")
        return None

# Funciones específicas para diferentes tipos de archivos

def save_photo_percha_cloudinary(percha_id, foto_uploaded):
    """Guardar foto de percha en Cloudinary"""
    return upload_photo_to_cloudinary(
        foto_uploaded,
        folder="fotos_perchas",
        public_id_prefix=f"percha_{percha_id}"
    )

def save_photo_comercial_cloudinary(entrega_id, foto_uploaded):
    """Guardar foto de entrega comercial en Cloudinary"""
    return upload_photo_to_cloudinary(
        foto_uploaded,
        folder="fotos_comerciales",
        public_id_prefix=f"comercial_{entrega_id}"
    )

def save_photo_letrero_cloudinary(letrero_id, foto_uploaded):
    """Guardar foto de letrero en Cloudinary"""
    return upload_photo_to_cloudinary(
        foto_uploaded,
        folder="fotos_letreros",
        public_id_prefix=f"letrero_{letrero_id}"
    )

def save_pdf_evento_cloudinary(pdf_path, evento_id, cliente):
    """Guardar PDF de autorización de evento en Cloudinary"""
    cliente_safe = cliente.replace("/", "_").replace("\\", "_")
    return upload_pdf_to_cloudinary(
        pdf_path,
        folder=f"documentos/eventos/{cliente_safe}",
        public_id_prefix=f"autorizacion_evento_{evento_id}"
    )

def save_pdf_letrero_cloudinary(pdf_path, letrero_id, cliente):
    """Guardar PDF de autorización de letrero en Cloudinary"""
    cliente_safe = cliente.replace("/", "_").replace("\\", "_")
    return upload_pdf_to_cloudinary(
        pdf_path,
        folder=f"documentos/letreros/{cliente_safe}",
        public_id_prefix=f"autorizacion_letrero_{letrero_id}"
    )

