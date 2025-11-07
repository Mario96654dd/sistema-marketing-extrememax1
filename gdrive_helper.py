# -*- coding: utf-8 -*-
"""Utilidades para almacenar archivos en Google Drive con cuentas de servicio."""

from __future__ import annotations

import json
import mimetypes
from dataclasses import dataclass
from io import BytesIO
from typing import Optional

import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload


_drive_service = None


def _get_credentials():
    secrets = st.secrets.get("gdrive", {})
    info_raw = secrets.get("service_account_json")
    if not info_raw:
        raise RuntimeError("No se encontró 'service_account_json' en los secrets de Streamlit.")
    if isinstance(info_raw, str):
        try:
            info = json.loads(info_raw)
        except json.JSONDecodeError:
            try:
                info = json.loads(info_raw, strict=False)
            except json.JSONDecodeError:
                sanitized = info_raw.replace("\r", "\\r").replace("\n", "\\n")
                info = json.loads(sanitized)
    else:
        info = info_raw
    credentials = service_account.Credentials.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return credentials


def _get_service():
    global _drive_service
    if _drive_service is None:
        credentials = _get_credentials()
        _drive_service = build("drive", "v3", credentials=credentials, cache_discovery=False)
    return _drive_service


def _ensure_folder(parent_id: str, folder_name: str) -> str:
    service = _get_service()
    query = (
        " and ".join([
            "mimeType='application/vnd.google-apps.folder'",
            "trashed=false",
            f"name='{folder_name}'",
            f"'{parent_id}' in parents"
        ])
    )
    results = service.files().list(
        q=query,
        fields="files(id, name)",
        pageSize=1
    ).execute()
    files = results.get("files", [])
    if files:
        return files[0]["id"]

    metadata = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id]
    }
    folder = service.files().create(body=metadata, fields="id").execute()
    return folder["id"]


def _guess_mimetype(filename: str, default: str = "application/octet-stream") -> str:
    mime, _ = mimetypes.guess_type(filename)
    return mime or default


@dataclass
class UploadResult:
    file_id: str
    url: str
    name: str


def _upload_bytes(data: bytes, filename: str, folder_path: list[str]) -> UploadResult:
    secrets = st.secrets.get("gdrive", {})
    root_id = secrets.get("folder_id")
    if not root_id:
        raise RuntimeError("No se encontró 'folder_id' en los secrets de Streamlit.")

    service = _get_service()

    current_parent = root_id
    for folder_name in folder_path:
        current_parent = _ensure_folder(current_parent, folder_name)

    media = MediaIoBaseUpload(BytesIO(data), mimetype=_guess_mimetype(filename), resumable=False)

    file_metadata = {
        "name": filename,
        "parents": [current_parent]
    }

    created = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, name"
    ).execute()

    file_id = created.get("id")
    if not file_id:
        raise RuntimeError("No se pudo obtener el ID del archivo subido a Drive.")

    # Obtener enlace web
    permissions = {
        "type": "anyone",
        "role": "reader"
    }
    try:
        service.permissions().create(fileId=file_id, body=permissions, allowFileDiscovery=False).execute()
    except Exception:
        # Si ya existe un permiso público, Google lanza error; lo ignoramos.
        pass

    file = service.files().get(fileId=file_id, fields="id, name").execute()
    url = f"https://drive.google.com/uc?export=view&id={file_id}"

    return UploadResult(
        file_id=file_id,
        url=url,
        name=file.get("name")
    )


def _read_file(file_obj) -> bytes:
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
    data = file_obj.read()
    if hasattr(file_obj, "seek"):
        try:
            file_obj.seek(0)
        except Exception:
            pass
    return data


def _sanitize_name(value: str) -> str:
    if not value:
        return "sin_nombre"
    return value.replace("/", "_").replace("\\", "_")


def save_photo_percha_drive(percha_id: str, file_obj) -> Optional[dict]:
    data = _read_file(file_obj)
    filename = getattr(file_obj, "name", None) or f"percha_{percha_id}.jpg"
    result = _upload_bytes(data, filename, ["perchas", _sanitize_name(str(percha_id))])
    return {"file_id": result.file_id, "url": result.url, "name": result.name}


def save_photo_comercial_drive(entrega_id: str, file_obj) -> Optional[dict]:
    data = _read_file(file_obj)
    filename = getattr(file_obj, "name", None) or f"comercial_{entrega_id}.jpg"
    result = _upload_bytes(data, filename, ["comerciales", _sanitize_name(str(entrega_id))])
    return {"file_id": result.file_id, "url": result.url, "name": result.name}


def save_photo_letrero_drive(letrero_id: str, file_obj) -> Optional[dict]:
    data = _read_file(file_obj)
    filename = getattr(file_obj, "name", None) or f"letrero_{letrero_id}.jpg"
    result = _upload_bytes(data, filename, ["letreros", _sanitize_name(str(letrero_id))])
    return {"file_id": result.file_id, "url": result.url, "name": result.name}


def save_pdf_evento_drive(pdf_path, evento_id: str, cliente: str) -> Optional[dict]:
    nombre = f"Autorizacion_Evento_{evento_id}.pdf"
    with open(pdf_path, "rb") as fh:
        data = fh.read()
    result = _upload_bytes(data, nombre, ["eventos", _sanitize_name(cliente)])
    return {"file_id": result.file_id, "url": result.url, "name": result.name}


def save_pdf_letrero_drive(pdf_path, letrero_id: str, cliente: str) -> Optional[dict]:
    nombre = f"Autorizacion_Letrero_{letrero_id}.pdf"
    with open(pdf_path, "rb") as fh:
        data = fh.read()
    result = _upload_bytes(data, nombre, ["letreros", _sanitize_name(cliente)])
    return {"file_id": result.file_id, "url": result.url, "name": result.name}


