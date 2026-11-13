from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from pathlib import Path, PurePosixPath
import os
import zipfile
import uuid
import shutil
import io
import subprocess
import tempfile
import re
import json
import threading
import time
from lxml import etree
import copy
from dotenv import load_dotenv
import requests
from datetime import datetime
import traceback

from pptx_geometry import SlideGeometry, absolute_rect, local_rect
from pptx_text import (
    apply_text_translations,
    extract_text_blocks as extract_pptx_text_blocks,
)

# Отключаем логи запросов Flask
import logging
logging.getLogger('werkzeug').setLevel(logging.ERROR)  # ← Только ошибки
# ИЛИ logging.getLogger('werkzeug').disabled = True     # ← Полностью отключить

# ========== ИНИЦИАЛИЗАЦИЯ ==========
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_UPLOAD_MB", "200")) * 1024 * 1024
_allowed_origins = [origin.strip() for origin in os.getenv(
    "CORS_ORIGINS", "http://127.0.0.1:5001,http://localhost:5001"
).split(",") if origin.strip()]
CORS(app, origins=_allowed_origins)

APP_HOST = os.getenv("APP_HOST", "127.0.0.1")
APP_PORT = int(os.getenv("APP_PORT", "5001"))
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1").rstrip("/")
DEEPSEEK_VERIFY_TLS = os.getenv("DEEPSEEK_VERIFY_TLS", "true").lower() not in {"0", "false", "no"}
PROMPT_BATCH_CHARS = max(4000, int(os.getenv("PROMPT_BATCH_CHARS", "24000")))

# ========== РАБОЧАЯ ДИРЕКТОРИЯ ==========
WORKSPACE_DIR = BASE_DIR / "workspace"
UPLOAD_DIR = WORKSPACE_DIR / "upload"
WORK_DIR = WORKSPACE_DIR / "work"
OUT_DIR = WORKSPACE_DIR / "out"
IMG_DIR = WORKSPACE_DIR / "img"
SESSIONS_DIR = WORKSPACE_DIR / "saved_sessions"
SESSIONS_DIR.mkdir(parents=True, exist_ok=True)

for d in [UPLOAD_DIR, WORK_DIR, OUT_DIR, IMG_DIR]:
    d.mkdir(parents=True, exist_ok=True)

SETTINGS_FILE = WORKSPACE_DIR / "settings.json"


def session_directory(folder_name):
    """Возвращает папку сессии, не допуская выхода за saved_sessions."""
    if not isinstance(folder_name, str) or not folder_name.strip():
        raise ValueError("Invalid session name")
    folder_name = folder_name.strip()
    if folder_name in {".", ".."} or Path(folder_name).name != folder_name:
        raise ValueError("Invalid session name")
    if (len(folder_name) > 160 or folder_name.startswith(".")
            or not re.fullmatch(r"[\w .()\[\]\-]+", folder_name, re.UNICODE)):
        raise ValueError("Invalid session name")
    return SESSIONS_DIR / folder_name

# ========== TXT ФАЙЛЫ ПРОМПТА ==========
PROMPT_ROLE_FILE = WORKSPACE_DIR / "prompt_role.txt"
PART1_FILE = WORKSPACE_DIR / "part1_do_not_translate.txt"
PART2_FILE = WORKSPACE_DIR / "part2_special_rules.txt"
PART3_FILE = WORKSPACE_DIR / "part3_technical.txt"
PART4_FILE = WORKSPACE_DIR / "part4_examples.txt"
GLOSSARY_TXT_FILE = WORKSPACE_DIR / "glossary.txt"
VARIANTS_FILE = WORKSPACE_DIR / "translation_variants.txt"

# ========== ГЛОБАЛЬНЫЕ ХРАНИЛИЩА ==========
sessions_info = {}
current_blocks_cache = []
updated_charts = {}
jobs = {}
EMU_PER_INCH = 914400

import threading
import time

def cleanup_old_sessions():
    """Очищает сессии старше 24 часов"""
    while True:
        time.sleep(3600)  # Проверка каждый час
        now = datetime.now()
        to_delete = []
        for sid, info in sessions_info.items():
            created = info.get('created_at')
            if created and (now - created).total_seconds() > 86400:  # 24 часа
                to_delete.append(sid)

        for sid in to_delete:
            print(f"🧹 Очистка старой сессии: {sid}")
            work_dir = Path(info['work_dir'])
            if work_dir.exists():
                shutil.rmtree(work_dir, ignore_errors=True)
            del sessions_info[sid]

# Запустить фоновую очистку
cleanup_thread = threading.Thread(target=cleanup_old_sessions, daemon=True)
cleanup_thread.start()

# ========== ПРОСТРАНСТВА ИМЕН PPTX ==========
NAMESPACES = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'
}
for prefix, uri in NAMESPACES.items():
    etree.register_namespace(prefix, uri)

@app.route('/api/save_translations/<session>', methods=['POST'])
def save_translations_to_session(session):
    """Сохраняет переводы в сессию на сервере."""
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404

    data = request.get_json()
    translations = data.get('translations', {})

    sessions_info[session]['translations'] = translations
    print(f"   💾 Переводы сохранены в сессию: {len(translations)} слайдов")

    return jsonify({"success": True})

@app.route('/api/save_selected_variant/<session>/<int:slide_idx>/<int:block_id>', methods=['POST'])
def save_selected_variant(session, slide_idx, block_id):
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404

    data = request.get_json()
    variant_id = data.get('variant_id')

    if 'selected_variants' not in sessions_info[session]:
        sessions_info[session]['selected_variants'] = {}
    if str(slide_idx) not in sessions_info[session]['selected_variants']:
        sessions_info[session]['selected_variants'][str(slide_idx)] = {}

    sessions_info[session]['selected_variants'][str(slide_idx)][str(block_id)] = variant_id
    return jsonify({"success": True})

@app.route('/api/get_selected_variants/<session>/<int:slide_idx>', methods=['GET'])
def get_selected_variants(session, slide_idx):
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404

    variants = sessions_info[session].get('selected_variants', {}).get(str(slide_idx), {})
    return jsonify({"selected_variants": variants})

@app.route('/api/markers/<session>/<int:slide_idx>', methods=['GET', 'POST'])
def handle_markers(session, slide_idx):
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404

    if request.method == 'GET':
        markers = sessions_info[session].get('markers', {}).get(str(slide_idx), {})
        return jsonify({"markers": markers})

    data = request.get_json()
    if 'markers' not in sessions_info[session]:
        sessions_info[session]['markers'] = {}
    sessions_info[session]['markers'][str(slide_idx)] = data.get('markers', {})
    return jsonify({"success": True})

@app.route('/api/areas/<session>/<int:slide_idx>', methods=['GET', 'POST'])
def handle_areas(session, slide_idx):
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404

    if request.method == 'GET':
        areas = sessions_info[session].get('areas', {}).get(str(slide_idx), [])
        return jsonify({"areas": areas})

    data = request.get_json()
    if 'areas' not in sessions_info[session]:
        sessions_info[session]['areas'] = {}
    sessions_info[session]['areas'][str(slide_idx)] = data.get('areas', [])
    return jsonify({"success": True})

@app.route('/api/workspace/clean', methods=['POST'])
def clean_workspace():
    """Очищает временные файлы workspace."""
    try:
        # Очищаем upload
        for f in UPLOAD_DIR.glob("*"):
            if f.is_file():
                f.unlink()

        # Очищаем work (кроме активных сессий)
        active_sessions = set(sessions_info.keys())
        for session_dir in WORK_DIR.iterdir():
            if session_dir.is_dir() and session_dir.name not in active_sessions:
                shutil.rmtree(session_dir)

        # Очищаем out
        for f in OUT_DIR.glob("*"):
            if f.is_file():
                f.unlink()

        # Очищаем img
        for f in IMG_DIR.glob("*"):
            if f.is_dir():
                shutil.rmtree(f)
            elif f.is_file():
                f.unlink()

        return jsonify({"success": True, "message": "Workspace cleaned"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def export_translations_to_txt(translations, total_slides, filepath):
    """Экспортирует переводы в TXT файл."""
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(f"# ПЕРЕВОДЫ СЕССИИ\n")
        f.write(f"# Сохранено: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"# Всего слайдов: {total_slides}\n\n")

        for slide_idx in range(total_slides):
            slide_trans = translations.get(str(slide_idx), translations.get(slide_idx, {}))
            if not slide_trans:
                continue

            f.write(f"\n{'='*50}\n")
            f.write(f"СЛАЙД {slide_idx + 1}\n")
            f.write(f"{'='*50}\n\n")

            for block_id, block_trans in slide_trans.items():
                if isinstance(block_trans, dict):
                    # Сортируем варианты
                    variant_keys = sorted(
                        [k for k in block_trans.keys() if k.endswith('_translation')],
                        key=lambda x: int(x.split('_')[0].replace('variant', '') or 1)
                    )

                    for v_idx, key in enumerate(variant_keys, 1):
                        value = block_trans[key]
                        if value:
                            quotes = "'" * v_idx
                            f.write(f"[{block_id}{quotes}] {value}\n")
                elif isinstance(block_trans, str):
                    f.write(f"[{block_id}'] {block_trans}\n")

            f.write("\n")

def list_saved_sessions():
    """Возвращает список сохраненных сессий."""
    sessions = []

    if not SESSIONS_DIR.exists():
        return sessions

    for folder in SESSIONS_DIR.iterdir():
        if not folder.is_dir():
            continue

        info_file = folder / "session_info.json"
        if not info_file.exists():
            continue

        try:
            with open(info_file, 'r', encoding='utf-8') as f:
                meta = json.load(f)

            # Проверяем наличие файлов
            work_exists = (folder / "work").is_dir()
            original_exists = any(folder.glob("*.pptx"))

            sessions.append({
                "folder": folder.name,
                "original_name": meta.get('original_name', ''),
                "saved_at": meta.get('saved_at', ''),
                "total_slides": meta.get('total_slides', 0),
                "has_translations": meta.get('has_translations', False),
                "work_exists": work_exists,
                "original_exists": original_exists,
                "meta": meta
            })
        except Exception as e:
            print(f"   ⚠️ Ошибка чтения {info_file}: {e}")

    # Сортируем по дате (новые сначала)
    sessions.sort(key=lambda x: x.get('saved_at', ''), reverse=True)
    return sessions


def load_saved_session(folder_name, new_session_id=None):
    """Загружает сохраненную сессию."""
    saved_dir = session_directory(folder_name)
    info_file = saved_dir / "session_info.json"

    if not info_file.exists():
        return {"success": False, "error": "Session info not found"}

    with open(info_file, 'r', encoding='utf-8') as f:
        meta = json.load(f)

    session_id = new_session_id or str(uuid.uuid4())

    # Копируем work_dir
    work_dir = WORK_DIR / session_id
    saved_work = saved_dir / "work"
    if saved_work.exists():
        if work_dir.exists():
            shutil.rmtree(work_dir)
        shutil.copytree(saved_work, work_dir)

    # Копируем оригинальный файл
    original_path = UPLOAD_DIR / f"{session_id}.pptx"
    saved_original = next(saved_dir.glob("*.pptx"), saved_dir / "missing.pptx")
    if saved_original.exists():
        shutil.copy2(saved_original, original_path)

    # ========== ЗАГРУЖАЕМ ПЕРЕВОДЫ ИЗ TXT ==========
    translations = {}
    translations_file = saved_dir / "translations.txt"
    translations_json = saved_dir / "translations.json"
    if translations_json.exists():
        translations = json.loads(translations_json.read_text(encoding="utf-8"))
    elif translations_file.exists():
        translations = import_translations_from_txt(translations_file)
    if translations:
        print(f"   📂 Переводы загружены для {len(translations)} слайдов")
    # =============================================

    # Восстанавливаем сессию
    sessions_info[session_id] = {
        'original_name': meta.get('original_name', ''),
        'original_path': str(original_path),
        'work_dir': str(work_dir),
        'total_slides': meta.get('total_slides', 0),
        'is_pdf': False,
        'translations': translations,
        'preview_ready': False,
        'preview_pages': 0,
        'created_at': datetime.now()
    }

    if 'markers' in meta:
        sessions_info[session_id]['markers'] = meta['markers']
    if 'areas' in meta:
        sessions_info[session_id]['areas'] = meta['areas']
    if 'selected_variants' in meta:
        sessions_info[session_id]['selected_variants'] = meta['selected_variants']
    updated_charts[session_id] = {}

    return {"success": True, "session_id": session_id, "meta": meta, "translations": {str(k): v for k, v in translations.items()}}


def import_translations_from_txt(filepath):
    """Импортирует переводы из TXT файла."""
    translations = {}

    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    current_slide = None

    for line in lines:
        line = line.strip()
        if not line or line.startswith('#'):
            continue

        # Определяем слайд
        slide_match = re.match(r'^СЛАЙД\s+(\d+)', line)
        if slide_match:
            current_slide = str(int(slide_match.group(1)) - 1)
            if current_slide not in translations:
                translations[current_slide] = {}
            continue

        # [block_id'] или [block_id''] или [block_id''']
        trans_match = re.match(r"^\[(\d+)('+)\]\s*(.*)", line)
        if trans_match and current_slide is not None:
            block_id = trans_match.group(1)
            quote_count = len(trans_match.group(2))
            translation = trans_match.group(3)

            if block_id not in translations[current_slide]:
                translations[current_slide][block_id] = {}

            translations[current_slide][block_id][f'variant_{quote_count}_translation'] = translation
            translations[current_slide][block_id][f'variant_{quote_count}_source'] = 'manual'

    return translations

@app.route('/api/sessions/save/<session>', methods=['POST'])
def api_save_session(session):
    try:
        data = request.get_json() or {}
        custom_name = data.get('name', None)
        overwrite = data.get('overwrite', False)
        result = save_session(session, custom_name, overwrite)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def save_session(session_id, custom_name=None, overwrite=False):
    """Сохраняет текущую сессию в папку saved_sessions.

    Args:
        session_id: ID сессии
        custom_name: пользовательское имя для сохранения
        overwrite: если True - перезаписывает существующую сессию с таким именем

    Returns:
        dict: результат операции
    """
    if session_id not in sessions_info:
        return {"success": False, "error": "Session not found"}

    info = sessions_info[session_id]
    original_name = info.get('original_name', 'presentation')
    date_str = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Определяем имя папки
    if overwrite and custom_name:
        # Перезапись существующей сессии
        folder_name = custom_name
        saved_dir = session_directory(folder_name)

        # Удаляем старую папку если существует
        if saved_dir.exists():
            shutil.rmtree(saved_dir)
            print(f"   🔄 Перезапись сессии: {folder_name}")
    else:
        # Обычное сохранение (новая сессия)
        if custom_name:
            folder_name = custom_name
        else:
            safe_original = re.sub(r"[^\w .()\[\]\-]+", "_", original_name, flags=re.UNICODE).strip(" .")
            folder_name = f"{safe_original or 'presentation'}_{date_str}"

    saved_dir = session_directory(folder_name)
    if saved_dir.exists() and not overwrite:
        return {"success": False, "error": "Session with this name already exists"}
    saved_dir.mkdir(parents=True, exist_ok=True)

    # Копируем work_dir
    work_dir = Path(info['work_dir'])
    dest_work = saved_dir / "work"
    if work_dir.exists():
        if dest_work.exists():
            shutil.rmtree(dest_work)
        shutil.copytree(work_dir, dest_work)

    # Копируем оригинальный файл
    original_path = Path(info['original_path'])
    dest_original = saved_dir / original_path.name
    shutil.copy2(original_path, dest_original)

    # JSON сохраняет выбранные стили и источники без потерь; TXT нужен для чтения.
    translations_file = saved_dir / "translations.txt"
    translations_json = saved_dir / "translations.json"
    if info.get('translations'):
        serializable = {str(k): v for k, v in info['translations'].items()}
        translations_json.write_text(
            json.dumps(serializable, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        export_translations_to_txt(serializable, info.get('total_slides', 0), translations_file)
        print(f"   💾 Переводы сохранены в {translations_file.name}")

    # Сохраняем метаданные
    meta = {
        "session_id": session_id,
        "original_name": original_name,
        "original_path": str(dest_original),
        "work_dir": str(dest_work),
        "total_slides": info.get('total_slides', 0),
        "saved_at": datetime.now().isoformat(),
        "custom_name": folder_name,
        "has_translations": bool(info.get('translations')),
        "translations_file": translations_file.name if translations_file.exists() else None,
        "overwritten": overwrite
    }

    # Сохраняем маркеры и области
    if 'markers' in info:
        meta['markers'] = info['markers']
    if 'areas' in info:
        meta['areas'] = info['areas']
    if 'selected_variants' in info:
        meta['selected_variants'] = info['selected_variants']

    with open(saved_dir / "session_info.json", 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

    print(f"   ✅ Сессия сохранена: {folder_name}")
    return {"success": True, "folder": folder_name, "path": str(saved_dir), "overwritten": overwrite}

@app.route('/api/sessions/list', methods=['GET'])
def api_list_sessions():
    """Возвращает список сохраненных сессий."""
    try:
        sessions = list_saved_sessions()
        return jsonify({"success": True, "sessions": sessions})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/sessions/load/<folder_name>', methods=['POST'])
def api_load_session(folder_name):
    """Загружает сохраненную сессию."""
    try:
        result = load_saved_session(folder_name)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/debug/slide_xml/<session>/<int:slide_idx>', methods=['GET'])
def debug_slide_xml(session, slide_idx):
    """Временный эндпоинт для отладки - возвращает XML слайда."""
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404

    info = sessions_info[session]
    work_dir = Path(info['work_dir'])
    slides = slides_list(work_dir)

    if slide_idx >= len(slides):
        return jsonify({"error": "Slide not found"}), 404

    xml_path = slides[slide_idx]

    try:
        with open(xml_path, 'r', encoding='utf-8') as f:
            xml_content = f.read()
        return xml_content, 200, {'Content-Type': 'text/xml'}
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/sessions/delete/<folder_name>', methods=['DELETE'])
def api_delete_session(folder_name):
    """Удаляет сохраненную сессию."""
    try:
        saved_dir = session_directory(folder_name)
        if saved_dir.exists():
            shutil.rmtree(saved_dir)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/status/<session>', methods=['GET'])
def get_session_status(session):
    """Возвращает текущий статус операций для сессии."""
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404

    info = sessions_info[session]

    status = {
        "session_active": True,
        "total_slides": info.get('total_slides', 0),
        "translations_loaded": len(info.get('translations', {})),
        "preview_ready": info.get('preview_ready', False),
        "last_operation": info.get('last_operation', ''),
        "last_update": info.get('last_update', '')
    }

    return jsonify(status)

# ============================================================
# 1. ЗАГРУЗКА TXT ФАЙЛОВ ПРОМПТА
# ============================================================

def load_prompt_role():
    """Загружает роль переводчика из prompt_role.txt"""
    if PROMPT_ROLE_FILE.exists():
        with open(PROMPT_ROLE_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return "Ты — профессиональный переводчик презентаций в области телекоммуникаций."


def load_part1_do_not_translate():
    """
    Загружает часть 1 (НЕ ПЕРЕВОДИТЬ) из TXT.
    Поддерживает два формата:
    1. Построчный:
       # Категория
       - элемент1
       - элемент2

    2. Через запятую:
       # Категория
       элемент1, элемент2, элемент3
    """
    if not PART1_FILE.exists():
        return []

    categories = []
    current_category = None

    with open(PART1_FILE, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue

            if line.startswith('#'):
                current_category = line[1:].strip()
                categories.append({'name': current_category, 'items': []})
            elif current_category is not None:
                if line.startswith('-'):
                    # Построчный формат
                    item = line[1:].strip()
                    if item:
                        categories[-1]['items'].append(item)
                else:
                    # Формат с запятыми - разбиваем всю строку
                    items = [item.strip() for item in line.split(',') if item.strip()]
                    categories[-1]['items'].extend(items)

    return categories

def load_part2_special_rules():
    """
    Загружает часть 2 (ОСОБЫЕ ПРАВИЛА) из TXT.
    Формат:
        # Название правила
        > Описание правила
        - исходный → перевод
    """
    if not PART2_FILE.exists():
        return []

    rules = []
    current_rule = None

    with open(PART2_FILE, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            if line.startswith('#'):
                if current_rule:
                    rules.append(current_rule)
                current_rule = {
                    'name': line[1:].strip(),
                    'description': '',
                    'examples': []
                }
            elif line.startswith('>') and current_rule:
                current_rule['description'] = line[1:].strip()
            elif line.startswith('-') and current_rule:
                example = line[1:].strip()
                if '→' in example:
                    source, target = example.split('→', 1)
                    current_rule['examples'].append({
                        'source': source.strip(),
                        'target': target.strip()
                    })
        if current_rule:
            rules.append(current_rule)
    return rules

def load_part3_technical():
    """Загружает часть 3 (ТЕХНИЧЕСКИЕ ТРЕБОВАНИЯ) из TXT."""
    if not PART3_FILE.exists():
        return []

    requirements = []
    with open(PART3_FILE, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line:
                if line.startswith('-'):
                    line = line[1:].strip()
                requirements.append(line)
    return requirements

def load_part4_examples():
    """Загружает примеры эталонных переводов."""
    if not PART4_FILE.exists():
        return []

    examples = []
    with open(PART4_FILE, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            if '→' in line:
                source, target = line.split('→', 1)
                examples.append({
                    'source': source.strip(),
                    'target': target.strip()
                })
    return examples

def load_glossary():
    """
    Загружает глоссарий из TXT.
    Формат: исходный → перевод
    """
    if not GLOSSARY_TXT_FILE.exists():
        return []

    terms = []
    with open(GLOSSARY_TXT_FILE, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            if '→' in line:
                source, target = line.split('→', 1)
                terms.append({
                    'source': source.strip(),
                    'target': target.strip()
                })
    return terms




# ============================================================
# 2. СОХРАНЕНИЕ TXT ФАЙЛОВ
# ============================================================

def save_glossary_txt(terms):
    """Сохраняет глоссарий в TXT."""
    with open(GLOSSARY_TXT_FILE, 'w', encoding='utf-8') as f:
        f.write("# Глоссарий\n# Формат: исходный_термин → перевод\n\n")
        for term in terms:
            f.write(f"{term['source']} → {term['target']}\n")


def save_part1_txt(categories):
    """Сохраняет часть 1 в TXT."""
    with open(PART1_FILE, 'w', encoding='utf-8') as f:
        for cat in categories:
            f.write(f"# {cat['name']}\n")
            for item in cat['items']:
                f.write(f"- {item}\n")
            f.write("\n")


def save_part2_txt(rules):
    """Сохраняет часть 2 в TXT."""
    with open(PART2_FILE, 'w', encoding='utf-8') as f:
        for rule in rules:
            f.write(f"# {rule['name']}\n")
            f.write(f"> {rule['description']}\n")
            for ex in rule.get('examples', []):
                f.write(f"- {ex['source']} → {ex['target']}\n")
            f.write("\n")


def save_part3_txt(requirements):
    """Сохраняет часть 3 в TXT."""
    with open(PART3_FILE, 'w', encoding='utf-8') as f:
        for req in requirements:
            f.write(f"- {req}\n")


# ============================================================
# 3. РАБОТА С ВАРИАНТАМИ ПЕРЕВОДА
# ============================================================

def load_translation_variants():
    """
    Загружает варианты из TXT.
    Формат:
        [ORIGINAL]
        Example1: текст
        Example2: текст

        [VARIANT]
        Code: [N']
        Name: Базовый
        Rule: правило
        Length: ±20%
        Example1: перевод
        Example2: перевод

        ---SEPARATOR---
    """
    if not VARIANTS_FILE.exists():
        return _get_default_variants_blocks()

    with open(VARIANTS_FILE, 'r', encoding='utf-8') as f:
        content = f.read()

    blocks = []
    for raw_block in content.split('---SEPARATOR---'):
        raw_block = raw_block.strip()
        if not raw_block:
            continue

        block = _parse_variant_block(raw_block)
        if block:
            blocks.append(block)

    return blocks if blocks else _get_default_variants_blocks()


def _parse_variant_block(raw_block):
    """Парсит один блок вариантов."""
    lines = raw_block.split('\n')
    originals = []
    variants = []
    current_variant = None
    current_section = None

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue

        if line.upper().startswith('[ORIGINAL]'):
            current_section = 'ORIGINAL'
        elif line.upper().startswith('[VARIANT'):
            if current_variant:
                variants.append(current_variant)
            current_section = 'VARIANT'
            current_variant = {'code': '', 'name': '', 'rule': '', 'length': '', 'examples': []}
        elif current_section == 'ORIGINAL' and line.lower().startswith('example'):
            parts = line.split(':', 1)
            if len(parts) == 2:
                originals.append(parts[1].strip())
        elif current_section == 'VARIANT' and current_variant is not None:
            if line.lower().startswith('code:'):
                current_variant['code'] = line.split(':', 1)[1].strip()
            elif line.lower().startswith('name:'):
                current_variant['name'] = line.split(':', 1)[1].strip()
            elif line.lower().startswith('rule:'):
                current_variant['rule'] = line.split(':', 1)[1].strip()
            elif line.lower().startswith('length:'):
                current_variant['length'] = line.split(':', 1)[1].strip()
            elif line.lower().startswith('example'):
                parts = line.split(':', 1)
                if len(parts) == 2:
                    current_variant['examples'].append(parts[1].strip())
        i += 1

    if current_variant:
        variants.append(current_variant)

    if not originals or not variants:
        return None

    return {'originals': originals, 'variants': variants}


def _get_default_variants_blocks():
    """Варианты по умолчанию."""
    return [{
        'originals': [
            'AI-powered Mobile Networks Unlock the Value of Connectivity',
            '5G networks enable new business models for operators'
        ],
        'variants': [
            {'code': "[N']", 'name': 'Базовый', 'rule': 'Точный перевод', 'length': '±20%',
             'examples': ['Мобильные сети на базе ИИ раскрывают ценность подключения',
                         'Сети 5G обеспечивают новые бизнес-модели для операторов']},
            {'code': "[N'']", 'name': 'Короткий', 'rule': 'Максимальное сжатие', 'length': 'короче на 30–50%',
             'examples': ['ИИ-сети раскрывают ценность подключений', '5G — новые бизнес-модели']},
            {'code': "[N''']", 'name': 'Деловой', 'rule': 'Бизнес-лексика', 'length': '±20%',
             'examples': ['Мобильные сети на ИИ позволяют монетизировать подключение',
                         '5G открывает операторам доступ к новым бизнес-моделям']},
            {'code': "[N'''']", 'name': 'Динамичный', 'rule': 'Активные глаголы', 'length': '±20%',
             'examples': ['Мобильные сети на базе ИИ начинают извлекать ценность из подключения',
                         'Благодаря 5G операторы могут внедрять новые бизнес-модели']},
            {'code': "[N''''']", 'name': 'Акцентный', 'rule': 'Ключевая мысль в начало', 'length': '±20%',
             'examples': ['Ценность подключения раскрывается — благодаря мобильным сетям на ИИ',
                         'Новые бизнес-модели для операторов дают сети 5G']},
            {'code': "[N'''''']", 'name': 'Тезисный', 'rule': 'Одна короткая фраза', 'length': 'короче на 40–60%',
             'examples': ['ИИ превращает подключение в ценность', '5G открывает операторам новые бизнес-модели']}
        ]
    }]

@app.route('/api/debug/variants', methods=['GET'])
def debug_variants():
    """Отладочный эндпоинт для проверки загрузки вариантов."""
    try:
        blocks = load_translation_variants()
        variants_list = get_variants_list()

        return jsonify({
            "blocks_count": len(blocks),
            "first_block_originals": blocks[0]['originals'] if blocks else [],
            "first_block_variants": blocks[0]['variants'] if blocks else [],
            "variants_list": variants_list
        })
    except Exception as e:
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500

def get_variants_list():
    """Возвращает плоский список вариантов для использования в интерфейсе."""
    blocks = load_translation_variants()
    if not blocks:
        return []

    first_block = blocks[0]
    variants = []

    for v in first_block['variants']:
        code = v.get('code', '')
        quote_count = code.count("'")
        variants.append({
            'id': f"variant_{quote_count}",
            'code': code,
            'name': v.get('name', ''),
            'description': v.get('rule', ''),
            'length': v.get('length', ''),
            'enabled': True
        })

    return variants


def save_translation_variants(blocks):
    """Сохраняет варианты в TXT."""
    with open(VARIANTS_FILE, 'w', encoding='utf-8') as f:
        for idx, block in enumerate(blocks):
            f.write("[ORIGINAL]\n")
            for i, ex in enumerate(block['originals'], 1):
                f.write(f"Example{i}: {ex}\n")
            f.write("\n")
            for variant in block['variants']:
                f.write("[VARIANT]\n")
                f.write(f"Code: {variant['code']}\n")
                f.write(f"Name: {variant['name']}\n")
                f.write(f"Rule: {variant['rule']}\n")
                f.write(f"Length: {variant['length']}\n")
                for i, ex in enumerate(variant['examples'], 1):
                    f.write(f"Example{i}: {ex}\n")
                f.write("\n")
            if idx < len(blocks) - 1:
                f.write("---SEPARATOR---\n\n")


# ============================================================
# 4. НАСТРОЙКИ (JSON)
# ============================================================

def load_settings():
    default = {
        "api_key": DEEPSEEK_API_KEY,
        "model": "deepseek-reasoner",
        "proxy_url": "",
        "show_separators": True,
        "sort_direction": "top-bottom",
        "blur": 0,
        "white_overlay": 0,
        "translation_priority": "api",
        "auto_topic": True,
        "use_compact": False
    }
    if SETTINGS_FILE.exists():
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
                default.update(saved)
        except:
            pass
    return default


def save_settings(settings):
    with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


USER_SETTINGS = load_settings()


# ============================================================
# 5. LIBREOFFICE
# ============================================================

def resolve_libreoffice():
    for candidate in ["soffice", "soffice.exe", "libreoffice",
                      r"C:\Program Files\LibreOffice\program\soffice.exe",
                      r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]:
        if isinstance(candidate, Path):
            if candidate.exists():
                return str(candidate)
        else:
            resolved = shutil.which(candidate)
            if resolved:
                return resolved
    return None


LIBREOFFICE_CMD = resolve_libreoffice()
LIBREOFFICE_AVAILABLE = LIBREOFFICE_CMD is not None
print(f"LibreOffice: {'✅ доступен' if LIBREOFFICE_AVAILABLE else '❌ не найден'}")


# ============================================================
# 6. PPTX: ИЗВЛЕЧЕНИЕ ТЕКСТА
# ============================================================

def unzip(path, dest):
    """Безопасно распаковывает OOXML-пакет с ограничением размера."""
    dest = Path(dest)
    dest.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "r") as archive:
        members = archive.infolist()
        if len(members) > 20000:
            raise ValueError("PPTX contains too many files")
        if sum(member.file_size for member in members) > 1024 * 1024 * 1024:
            raise ValueError("Unpacked PPTX is too large")
        for member in members:
            normalized = member.filename.replace("\\", "/")
            relative = PurePosixPath(normalized)
            if relative.is_absolute() or ".." in relative.parts or member.flag_bits & 1:
                raise ValueError("Unsafe or encrypted PPTX entry")
            target = (dest / Path(*relative.parts)).resolve()
            if dest.resolve() not in target.parents and target != dest.resolve():
                raise ValueError("Unsafe PPTX path")
        archive.extractall(dest)

def zipdir(folder, out):
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(folder):
            for f in files:
                full = os.path.join(root, f)
                rel = os.path.relpath(full, folder)
                z.write(full, rel)

def slides_list(work_dir):
    p = os.path.join(work_dir, "ppt/slides")
    if not os.path.exists(p):
        return []
    files = sorted([f for f in os.listdir(p) if f.startswith("slide") and f.endswith(".xml")],
                   key=lambda x: int(x.replace("slide", "").replace(".xml", "")))
    return [os.path.join(p, f) for f in files]

def get_slide_size(work_dir):
    pres_path = os.path.join(work_dir, "ppt/presentation.xml")
    if not os.path.exists(pres_path):
        return 10, 7.5
    try:
        tree = etree.parse(pres_path)
        sldSz = tree.find('.//p:sldSz', namespaces=NAMESPACES)
        if sldSz is not None:
            return int(sldSz.get('cx')) / EMU_PER_INCH, int(sldSz.get('cy')) / EMU_PER_INCH
    except:
        pass
    return 10, 7.5

# ============================================================
# КООРДИНАТЫ БЛОКОВ С УЧЕТОМ ГРУПП
# ============================================================
# Математика вынесена в pptx_geometry.py. Эти небольшие обёртки
# сохранены для совместимости с диагностическими API.

def set_slide_size_emu(width_emu, height_emu):
    """Больше не требуется: результат не подбирается по размеру слайда."""


def clear_coordinate_cache():
    """Геометрия не использует глобальный кэш."""


def _get_local_coordinates(element):
    rect = local_rect(element)
    return (int(rect.x), int(rect.y)) if rect else (0, 0)


def _get_element_size(element):
    rect = absolute_rect(element, include_own_transform=False)
    return (int(rect.width), int(rect.height)) if rect else (0, 0)


def get_absolute_coordinates(element):
    rect = absolute_rect(element)
    return (rect.x, rect.y) if rect else (0.0, 0.0)

# ============================================================
# ИЗВЛЕЧЕНИЕ ТЕКСТА, ТАБЛИЦ И ДИАГРАММ
# ============================================================

def extract_text_blocks(xml_path, pptx_zip=None, slide_rel_path=None):
    return extract_pptx_text_blocks(xml_path, pptx_zip, slide_rel_path)


def get_all_parent_groups(element):
    """
    Возвращает список всех родительских групп для отладки.
    """
    groups = []
    current = element.getparent()
    depth = 0

    while current is not None and depth < 20:
        if current.tag.endswith('grpSp'):
            # Получаем название группы, если есть
            nvPr = current.find('.//p:nvPr', namespaces=NAMESPACES)
            name = ""
            if nvPr is not None:
                cNvPr = nvPr.find('.//p:cNvPr', namespaces=NAMESPACES)
                if cNvPr is not None:
                    name = cNvPr.get('name', '')

            # Получаем смещение группы
            xfrm = current.find('.//a:xfrm', namespaces=NAMESPACES)
            off_x, off_y = 0, 0
            if xfrm is not None:
                off = xfrm.find('.//a:off', namespaces=NAMESPACES)
                if off is not None:
                    off_x = int(off.get('x', 0))
                    off_y = int(off.get('y', 0))

            groups.append({
                'depth': depth,
                'name': name,
                'offset_x': off_x,
                'offset_y': off_y,
                'offset_inches': (off_x / EMU_PER_INCH, off_y / EMU_PER_INCH)
            })

        current = current.getparent()
        depth += 1

    return groups

@app.route('/api/debug/coordinates/<session>/<int:slide_idx>', methods=['GET'])
def debug_slide_coordinates(session, slide_idx):
    """
    Отладочный эндпоинт для проверки координат блоков.
    Возвращает подробную информацию о каждом блоке и его родительских группах.
    """
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404

    info = sessions_info[session]
    work_dir = Path(info['work_dir'])

    # Получаем список всех слайдов
    slides = slides_list(work_dir)

    if slide_idx >= len(slides):
        return jsonify({"error": f"Slide {slide_idx} not found. Total slides: {len(slides)}"}), 404

    # Получаем размер слайда
    slide_width_inch, slide_height_inch = get_slide_size(work_dir)
    slide_width_emu = int(slide_width_inch * EMU_PER_INCH)
    slide_height_emu = int(slide_height_inch * EMU_PER_INCH)

    # Устанавливаем размер слайда
    set_slide_size_emu(slide_width_emu, slide_height_emu)

    # Очищаем кэш координат
    clear_coordinate_cache()

    slide_xml_path = slides[slide_idx]
    slide_rel_path = None
    rel_file_path = slide_xml_path.replace('/slides/', '/slides/_rels/') + '.rels'

    debug_data = {
        "session": session,
        "slide_idx": slide_idx,
        "slide_num": slide_idx + 1,
        "slide_width_inch": slide_width_inch,
        "slide_height_inch": slide_height_inch,
        "slide_width_emu": slide_width_emu,
        "slide_height_emu": slide_height_emu,
        "blocks": [],
        "groups_info": []
    }

    try:
        with zipfile.ZipFile(info['original_path'], 'r') as zf:
            # Проверяем .rels файл
            try:
                zf.getinfo(rel_file_path)
                slide_rel_path = rel_file_path
                debug_data["rels_file_exists"] = True
                debug_data["rels_file_path"] = rel_file_path
            except KeyError:
                debug_data["rels_file_exists"] = False
                debug_data["rels_file_path"] = None

            # Парсим XML слайда
            try:
                with zf.open(slide_xml_path) as f:
                    slide_tree = etree.fromstring(f.read())
            except Exception as e:
                return jsonify({"error": f"Failed to parse slide XML: {str(e)}"}), 500

            # Извлекаем блоки с отладочной информацией
            blocks = extract_text_blocks(slide_xml_path, zf, slide_rel_path)

            # Для каждого блока собираем дополнительную отладочную информацию
            containers = slide_tree.xpath('//p:sp | //a:tbl | //p:graphicFrame', namespaces=NAMESPACES)

            for idx, container in enumerate(containers[:len(blocks)]):  # Берем столько же, сколько блоков
                if idx >= len(blocks):
                    break

                block = blocks[idx]

                # Получаем информацию о родительских группах
                parent_groups = get_all_parent_groups(container)

                # Получаем локальные координаты (до применения групп)
                local_x, local_y = _get_local_coordinates(container)
                local_w, local_h = _get_element_size(container)

                # Проверяем, находится ли блок в пределах слайда
                x_in_slide = 0 <= block['x'] * EMU_PER_INCH <= slide_width_emu
                y_in_slide = 0 <= block['y'] * EMU_PER_INCH <= slide_height_emu

                debug_data["blocks"].append({
                    "id": block['id'],
                    "type": block.get('type', 'shape'),
                    "text_preview": block['text'][:100] + "..." if len(block['text']) > 100 else block['text'],
                    "text_length": len(block['text']),
                    "runs_count": block.get('runs_count', 1),
                    "local_coordinates_emu": {
                        "x": local_x,
                        "y": local_y,
                        "width": local_w,
                        "height": local_h
                    },
                    "local_coordinates_inch": {
                        "x": local_x / EMU_PER_INCH,
                        "y": local_y / EMU_PER_INCH,
                        "width": local_w / EMU_PER_INCH,
                        "height": local_h / EMU_PER_INCH
                    },
                    "absolute_coordinates_emu": {
                        "x": block['x'] * EMU_PER_INCH,
                        "y": block['y'] * EMU_PER_INCH,
                        "width": block.get('w', 0) * EMU_PER_INCH,
                        "height": block.get('h', 0) * EMU_PER_INCH
                    },
                    "absolute_coordinates_inch": {
                        "x": block['x'],
                        "y": block['y'],
                        "width": block.get('w', 0),
                        "height": block.get('h', 0)
                    },
                    "inside_slide": x_in_slide and y_in_slide,
                    "parent_groups_count": len(parent_groups),
                    "parent_groups": parent_groups
                })

            debug_data["total_blocks"] = len(blocks)
            debug_data["blocks_inside_slide"] = sum(1 for b in debug_data["blocks"] if b["inside_slide"])
            debug_data["blocks_outside_slide"] = debug_data["total_blocks"] - debug_data["blocks_inside_slide"]

    except Exception as e:
        print(f"   ❌ Ошибка отладки слайда {slide_idx}: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Debug error: {str(e)}"}), 500

    return jsonify(debug_data)


@app.route('/api/debug/coordinates/all/<session>', methods=['GET'])
def debug_all_slides_coordinates(session):
    """
    Отладочный эндпоинт для проверки координат на всех слайдах.
    """
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404

    info = sessions_info[session]
    work_dir = Path(info['work_dir'])

    slides = slides_list(work_dir)
    total_slides = len(slides)

    results = []

    for slide_idx in range(total_slides):
        try:
            # Получаем размер слайда
            slide_width_inch, slide_height_inch = get_slide_size(work_dir)
            slide_width_emu = int(slide_width_inch * EMU_PER_INCH)
            slide_height_emu = int(slide_height_inch * EMU_PER_INCH)

            set_slide_size_emu(slide_width_emu, slide_height_emu)
            clear_coordinate_cache()

            slide_xml_path = slides[slide_idx]

            with zipfile.ZipFile(info['original_path'], 'r') as zf:
                blocks = extract_text_blocks(slide_xml_path, zf, None)

            blocks_inside = 0
            blocks_outside = 0

            for block in blocks:
                x_emu = block['x'] * EMU_PER_INCH
                y_emu = block['y'] * EMU_PER_INCH
                if 0 <= x_emu <= slide_width_emu and 0 <= y_emu <= slide_height_emu:
                    blocks_inside += 1
                else:
                    blocks_outside += 1

            results.append({
                "slide_idx": slide_idx,
                "slide_num": slide_idx + 1,
                "total_blocks": len(blocks),
                "blocks_inside": blocks_inside,
                "blocks_outside": blocks_outside,
                "status": "ok" if blocks_outside == 0 else "warning"
            })

        except Exception as e:
            results.append({
                "slide_idx": slide_idx,
                "slide_num": slide_idx + 1,
                "error": str(e),
                "status": "error"
            })

    return jsonify({
        "session": session,
        "total_slides": total_slides,
        "slides": results,
        "summary": {
            "total_blocks": sum(r.get("total_blocks", 0) for r in results),
            "total_blocks_inside": sum(r.get("blocks_inside", 0) for r in results),
            "total_blocks_outside": sum(r.get("blocks_outside", 0) for r in results),
            "slides_with_errors": sum(1 for r in results if r.get("status") == "error"),
            "slides_with_outside_blocks": sum(1 for r in results if r.get("blocks_outside", 0) > 0)
        }
    })

# ============================================================
# 7. PPTX: ПРИМЕНЕНИЕ ПЕРЕВОДА
# ============================================================

def apply_translation_safe(xml_path, order, translations_list, blocks_cache, session):
    """Применяет выбранные переводы по стабильным ID блоков."""
    translations_by_id = {}
    for index, block_id in enumerate(order):
        if index >= len(translations_list):
            continue
        value = translations_list[index]
        if isinstance(value, dict):
            value = value.get("translation") or value.get("variant_1_translation") or ""
        if isinstance(value, str) and value.strip():
            translations_by_id[block_id] = value

    if not translations_by_id:
        return {"success": False, "applied_blocks": 0, "errors": [], "error": "Нет переводов"}

    chart_cache = updated_charts.setdefault(session, {})
    report = apply_text_translations(
        xml_path,
        translations_by_id,
        blocks_cache,
        sessions_info[session]["original_path"],
        chart_cache,
    )
    if report["applied_blocks"]:
        # Диаграммы сразу попадают в рабочий пакет, поэтому сохранение сессии
        # не теряет их даже до первого экспорта PPTX.
        work_dir = Path(sessions_info[session]["work_dir"])
        for chart_file, chart_tree in chart_cache.items():
            chart_path = work_dir / chart_file
            chart_path.parent.mkdir(parents=True, exist_ok=True)
            temporary = chart_path.with_suffix(chart_path.suffix + ".tmp")
            chart_tree.getroottree().write(
                str(temporary), encoding="utf-8", xml_declaration=True,
                standalone="yes", pretty_print=False,
            )
            os.replace(temporary, chart_path)
    return report


# ============================================================
# 8. PPTX: ЭКСПОРТ
# ============================================================

@app.route('/api/export/<session>', methods=['GET'])
def export(session):
    try:
        if session not in sessions_info:
            return jsonify({"error": "Session not found"}), 404

        info = sessions_info[session]
        work_dir = Path(info['work_dir'])
        original_name = info['original_name']

        temp_dir = Path(OUT_DIR) / session
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
        temp_dir.mkdir(parents=True, exist_ok=True)

        for item in work_dir.iterdir():
            dest = temp_dir / item.name
            if item.is_dir():
                shutil.copytree(item, dest)
            else:
                shutil.copy2(item, dest)

        # Удаляем preview
        preview_dir = temp_dir / "preview"
        if preview_dir.exists():
            shutil.rmtree(preview_dir)

        # Сохраняем диаграммы
        if session in updated_charts:
            for chart_file, chart_tree in updated_charts[session].items():
                chart_path = temp_dir / chart_file
                chart_path.parent.mkdir(parents=True, exist_ok=True)
                chart_tree.getroottree().write(
                    str(chart_path), encoding='utf-8', xml_declaration=True,
                    standalone='yes', pretty_print=False,
                )

        output_path = Path(OUT_DIR) / f"{session}.pptx"
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
            for root_dir, dirs, files in os.walk(temp_dir):
                if Path(root_dir) == temp_dir and '[Content_Types].xml' in files:
                    files.remove('[Content_Types].xml')
                    files.insert(0, '[Content_Types].xml')
                for file in files:
                    file_path = Path(root_dir) / file
                    zf.write(file_path, file_path.relative_to(temp_dir))

        shutil.rmtree(temp_dir)

        return send_file(output_path, as_attachment=True,
                        download_name=f"{original_name}_translated.pptx",
                        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')
    except Exception as e:
        print(f"❌ Ошибка экспорта: {e}")
        return jsonify({"error": str(e)}), 500


# ============================================================
# 9. DEEPSEEK API
# ============================================================


def _build_all_slides_prompt(all_texts_with_meta, target_lang, enabled_variant_ids=None):
    """Формирует промпт с запросом нескольких вариантов."""

    lang_name = {'ru': 'русский', 'en': 'английский'}.get(target_lang, target_lang)
    role = load_prompt_role()
    part1 = load_part1_do_not_translate()
    part2 = load_part2_special_rules()
    part3 = load_part3_technical()
    glossary = load_glossary()

    total_blocks = len(all_texts_with_meta)
    total_slides = len(set(b['slide_num'] for b in all_texts_with_meta))

    # Получаем все варианты из TXT
    all_variants = get_variants_list()

    # Определяем активные варианты (сохраняя исходный порядок из TXT)
    if enabled_variant_ids:
        # Фильтруем, сохраняя порядок из TXT
        active_variants = []
        for v in all_variants:
            if v['id'] in enabled_variant_ids:
                active_variants.append(v)
    else:
        # По умолчанию — все включенные
        active_variants = [v for v in all_variants if v.get('enabled', True)]

    variants_count = len(active_variants)

    # Если нет активных вариантов — используем только базовый
    if variants_count == 0:
        active_variants = [all_variants[0]] if all_variants else [{'id': 'variant_1', 'name': 'Базовый', 'code': "[N']", 'description': ''}]
        variants_count = 1

    prompt = role + "\n\n"
    prompt += f"Переведи {total_blocks} текстов с {total_slides} слайдов на {lang_name} язык.\n"
    prompt += f"Для каждого блока нужно предоставить {variants_count} вариантов перевода.\n\n"

    # Часть 1: НЕ ПЕРЕВОДИТЬ
    if part1:
        prompt += "=" * 50 + "\nЧАСТЬ 1. НЕ ПЕРЕВОДИ:\n" + "=" * 50 + "\n"
        for cat in part1:
            prompt += f"  • {cat['name']}: {', '.join(cat['items'])}\n"
        prompt += "\n"

    # Часть 2: ОСОБЫЕ ПРАВИЛА
    if part2:
        prompt += "=" * 50 + "\nЧАСТЬ 2. ОСОБЫЕ ПРАВИЛА:\n" + "=" * 50 + "\n"
        for rule in part2:
            prompt += f"  • {rule['name']}: {rule['description']}\n"
            for ex in rule.get('examples', []):
                prompt += f"    {ex['source']} → {ex['target']}\n"
        prompt += "\n"

    # Часть 3: ТЕХНИЧЕСКИЕ ТРЕБОВАНИЯ
    if part3:
        prompt += "=" * 50 + "\nЧАСТЬ 3. ТЕХНИЧЕСКИЕ ТРЕБОВАНИЯ:\n" + "=" * 50 + "\n"
        for i, req in enumerate(part3, 1):
            prompt += f"  {i}. {req}\n"
        prompt += "\n"

    # ГЛОССАРИЙ
    if glossary:
        prompt += "=" * 50 + "\nГЛОССАРИЙ:\n" + "=" * 50 + "\n"
        for term in glossary:
            prompt += f"  • {term['source']} → {term['target']}\n"
        prompt += "\n"

    # ПРИМЕРЫ СТИЛЕЙ (из TXT)
    variants_block = load_translation_variants()
    if variants_block:
        fb = variants_block[0]
        prompt += "=" * 50 + "\nПРИМЕРЫ СТИЛЕЙ:\n" + "=" * 50 + "\n"
        prompt += "Оригиналы:\n"
        for i, ex in enumerate(fb['originals'], 1):
            prompt += f"  {i}. {ex}\n"
        prompt += "\n"
        for v in fb['variants']:
            prompt += f"--- {v['name']} ({v['code']}) ---\n"
            prompt += f"Правило: {v['rule']}\nДлина: {v['length']}\nПереводы:\n"
            for i, ex in enumerate(v['examples'], 1):
                prompt += f"  {i}. {ex}\n"
            prompt += "\n"

    # ФОРМАТ ОТВЕТА
    prompt += "=" * 50 + "\n"
    prompt += "ФОРМАТ ОТВЕТА:\n"
    prompt += "=" * 50 + "\n\n"

    prompt += f"""Верни СТРОГО JSON следующего формата:

{{
  "translations": [
"""

    # Пример для первого блока
    example_variants = [f'"{v["name"]}_вариант_блока1"' for v in active_variants]
    prompt += f"    [{', '.join(example_variants)}],\n"
    prompt += f"    [{', '.join([f'"{v["name"]}_вариант_блока2"' for v in active_variants])}],\n"
    prompt += f"    ...\n"
    prompt += f"  ]\n}}\n\n"

    prompt += f"Где:\n"
    prompt += f"- translations — массив из {total_blocks} элементов\n"
    prompt += f"- Каждый элемент — массив из {variants_count} вариантов перевода\n"
    prompt += f"- Порядок вариантов внутри блока строго соответствует списку:\n"

    for i, v in enumerate(active_variants, 1):
        prompt += f"  {i}. {v['name']} ({v['code']}) — {v['description']}\n"

    prompt += f"\nВсего блоков: {total_blocks}, вариантов на каждый блок: {variants_count}\n\n"

    prompt += "=" * 50 + "\n"
    prompt += "=" * 50 + "\n"
    prompt += "ПРАВИЛО ДЛЯ КОРОТКИХ БЛОКОВ (менее 30 символов):\n"
    prompt += "=" * 50 + "\n\n"
    prompt += f"Для блоков короче 30 символов предоставляй ТОЛЬКО ОДИН перевод — в стиле {active_variants[0]['name'] if active_variants else 'Базовый'}.\n"
    prompt += "Для остальных стилей возвращай пустые строки.\n"
    prompt += "Пример:\n"
    prompt += "- Исходный блок: '3 hours | Qatar O' (менее 30 символов)\n"
    prompt += f'- Ответ: ["3 часа | Катар O", "", "", ""]  (если выбрано 4 стиля)\n'
    prompt += "\n"
    prompt += "- Исходный блок: 'Bubble size: | Outage occurrence duration (hour)' (более 30 символов)\n"
    prompt += "- Ответ: [полный набор вариантов для всех стилей]\n\n"


    # ТЕКСТЫ
    prompt += "=" * 50 + f"\nТЕКСТЫ ({total_blocks}):\n" + "=" * 50 + "\n"
    current_slide = None
    for b in all_texts_with_meta:
        if b['slide_num'] != current_slide:
            current_slide = b['slide_num']
            prompt += f"\n─── СЛАЙД {current_slide} ───\n"
        prompt += f"[{b['block_num']}] {b['text']}\n"

    prompt += f"\n{'=' * 50}\n"
    prompt += f"ВАЖНО: Верни ТОЛЬКО JSON с {total_blocks} массивами по {variants_count} вариантов.\n"
    prompt += f"Не добавляй пояснений, markdown, или дополнительного текста.\n"
    prompt += "=" * 50

    return prompt, active_variants

# ============================================================
# 10. API ЭНДПОИНТЫ
# ============================================================
@app.route('/api/translate/slides_selected/<session>', methods=['POST'])
def translate_selected_slides(session):
    """Переводит выбранные слайды через общий проверенный цикл."""
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404
    try:
        data = request.get_json(silent=True) or {}
        return jsonify(_perform_translation(session, data.get("slides", []), data))
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        print(f"❌ Ошибка перевода выбранных слайдов: {exc}")
        traceback.print_exc()
        return jsonify({"error": str(exc)}), 500


@app.route('/api/export_prompt_selected/<session>', methods=['POST'])
def export_prompt_selected(session):
    """Экспорт промпта для выбранных слайдов."""
    try:
        if session not in sessions_info:
            return jsonify({"error": "Session not found"}), 404

        data = request.get_json()
        slides_to_include = data.get('slides', [])
        target_lang = data.get('target_lang', 'ru')
        enabled_variant_ids = data.get('enabled_variants', None)

        if not slides_to_include:
            return jsonify({"error": "No slides selected"}), 400

        print(f"   📊 Генерация промпта для {len(slides_to_include)} слайдов...")

        info = sessions_info[session]
        all_slides = slides_list(info['work_dir'])

        # Собираем тексты только с выбранных слайдов
        all_texts = []
        skipped_russian = 0
        skipped_untranslatable = 0

        with zipfile.ZipFile(info['original_path'], 'r') as zf:
            for slide_idx in slides_to_include:
                if slide_idx >= len(all_slides):
                    continue
                blocks = extract_text_blocks(all_slides[slide_idx], zf)

                for block in blocks:
                    text = block.get('text', '')
                    runs = block.get('runs', [])

                    # Пропускаем русские
                    if has_russian(text):
                        skipped_russian += 1
                        continue

                    # Пропускаем непереводимые
                    if should_skip_translation(text):
                        skipped_untranslatable += 1
                        continue

                    if len(runs) > 1:
                        text = ' | '.join([r.get('text', '') for r in runs])

                    all_texts.append({
                        'slide_num': slide_idx + 1,
                        'block_num': len(all_texts) + 1,
                        'block_id': block['id'],
                        'text': text,
                        'has_runs': len(runs) > 1,
                        'runs_count': len(runs)
                    })

        print(f"   📝 Собрано блоков: {len(all_texts)}")
        print(f"   ⏭️ Пропущено русских: {skipped_russian}")
        print(f"   ⏭️ Пропущено непереводимых: {skipped_untranslatable}")

        if not all_texts:
            return jsonify({"error": "No text blocks found"}), 400

        # Формируем промпт
        print(f"   🤖 Формирование промпта...")
        prompt, active_variants = _build_all_slides_prompt(all_texts, target_lang, enabled_variant_ids)
        print(f"   📏 Размер промпта: {len(prompt)} символов (~{len(prompt)//4} токенов)")
        print(f"   🎨 Активных вариантов: {len(active_variants)}")

        # Сохраняем во временный файл
        prompt_file = OUT_DIR / f"prompt_selected_{session}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        with open(prompt_file, 'w', encoding='utf-8') as f:
            f.write(prompt)

        print(f"   💾 Файл сохранен: {prompt_file}")

        return send_file(
            prompt_file,
            as_attachment=True,
            download_name=f"prompt_slides_{len(slides_to_include)}.txt",
            mimetype='text/plain'
        )

    except Exception as e:
        print(f"❌ Ошибка экспорта промпта: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route('/')
def index():
    html_path = BASE_DIR / 'index.html'
    if html_path.exists():
        return send_file(html_path)
    return "<h1>index.html не найден</h1>"


@app.route('/api/settings', methods=['GET', 'POST'])
def api_settings():
    if request.method == 'GET':
        return jsonify(USER_SETTINGS)
    data = request.get_json()
    if data:
        USER_SETTINGS.update(data)
        save_settings(USER_SETTINGS)
        return jsonify({"success": True})
    return jsonify({"error": "No data"}), 400


@app.route('/api/upload', methods=['POST'])
def upload():
    """Загружает PPTX файл, распаковывает и генерирует превью."""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'Empty filename'}), 400

        if Path(file.filename).suffix.lower() != ".pptx":
            return jsonify({"error": "Only .pptx files are supported"}), 415

        session = str(uuid.uuid4())
        original_name = Path(file.filename).stem[:160] or "presentation"
        pptx_path = UPLOAD_DIR / f"{session}.pptx"
        work_dir = WORK_DIR / session
        file.save(pptx_path)
        try:
            unzip(pptx_path, work_dir)
            if not (work_dir / "ppt" / "presentation.xml").is_file():
                raise ValueError("Invalid PPTX: presentation.xml is missing")
        except (zipfile.BadZipFile, ValueError) as exc:
            pptx_path.unlink(missing_ok=True)
            shutil.rmtree(work_dir, ignore_errors=True)
            return jsonify({"error": f"Invalid PPTX: {exc}"}), 400

        # Считаем количество слайдов
        slides = slides_list(work_dir)
        total_slides = len(slides)
        if not total_slides:
            pptx_path.unlink(missing_ok=True)
            shutil.rmtree(work_dir, ignore_errors=True)
            return jsonify({"error": "PPTX contains no slides"}), 400

        # Сохраняем информацию о сессии
        sessions_info[session] = {
            'original_name': original_name,
            'original_path': str(pptx_path),
            'work_dir': str(work_dir),
            'total_slides': total_slides,
            'is_pdf': False,
            'preview_ready': False,
            'preview_pages': 0,
            'created_at': datetime.now()
        }

        # Генерируем превью в фоновом потоке
        if LIBREOFFICE_AVAILABLE:
            def generate_previews_bg():
                try:
                    preview_dir = work_dir / "preview"
                    preview_dir.mkdir(parents=True, exist_ok=True)

                    print(f"   🔄 Запущена генерация превью для сессии {session[:8]}...")

                    # Конвертируем PPTX в PDF (увеличиваем таймаут до 300 секунд)
                    cmd = [LIBREOFFICE_CMD, "--headless", "--convert-to", "pdf", "--outdir", str(preview_dir), str(pptx_path)]
                    result = subprocess.run(cmd, capture_output=True, timeout=300)

                    if result.returncode != 0:
                        print(f"   ⚠️ LibreOffice error: {result.stderr.decode('utf-8', errors='ignore')[:500]}")
                        return

                    # Находим сгенерированный PDF
                    pdf_file = preview_dir / f"{pptx_path.stem}.pdf"
                    if not pdf_file.exists():
                        print(f"   ⚠️ PDF не создан: {pdf_file}")
                        return

                    print(f"   📄 PDF создан: {pdf_file}")

                    # Конвертируем PDF в PNG
                    try:
                        try:
                            import pymupdf as fitz
                        except ImportError:
                            import fitz
                        doc = fitz.open(str(pdf_file))

                        count = 0
                        for idx, page in enumerate(doc):
                            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                            out_path = preview_dir / f"slide_{idx + 1}.png"
                            pix.save(str(out_path))
                            count += 1
                        doc.close()

                        # Обновляем статус
                        sessions_info[session]['preview_ready'] = True
                        sessions_info[session]['preview_pages'] = count
                        print(f"   ✅ Сгенерировано {count} PNG превью для сессии {session[:8]}...")

                    except ImportError:
                        print("   ⚠️ PyMuPDF не установлен, превью недоступны")
                    except Exception as e:
                        print(f"   ⚠️ Ошибка конвертации PDF в PNG: {e}")
                        # Пробуем альтернативный метод
                        try:
                            try:
                                import pymupdf as fitz
                            except ImportError:
                                import fitz
                            doc = fitz.open(str(pdf_file))
                            count = 0
                            for idx, page in enumerate(doc):
                                pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                                out_path = preview_dir / f"slide_{idx + 1}.png"
                                pix.save(str(out_path))
                                count += 1
                            doc.close()
                            sessions_info[session]['preview_ready'] = True
                            sessions_info[session]['preview_pages'] = count
                            print(f"   ✅ Сгенерировано {count} PNG превью (альтернативный метод)")
                        except Exception as e2:
                            print(f"   ⚠️ Альтернативный метод тоже не сработал: {e2}")

                except subprocess.TimeoutExpired:
                    print(f"   ⚠️ Таймаут генерации PDF (300 секунд)")
                except Exception as e:
                    print(f"   ⚠️ Ошибка генерации превью: {e}")
                    import traceback
                    traceback.print_exc()

            # Запускаем в фоне
            threading.Thread(target=generate_previews_bg, daemon=True).start()
            print(f"   🔄 Фоновая генерация превью запущена")
        else:
            print("   ⚠️ LibreOffice не найден, превью недоступны")

        return jsonify({
            "session": session,
            "total_slides": total_slides,
            "original_name": original_name,
            "is_pdf": False
        })

    except Exception as e:
        print(f"❌ Ошибка загрузки: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route('/api/preview_status/<session>', methods=['GET'])
def preview_status(session):
    """Возвращает детальный статус генерации превью."""
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404

    info = sessions_info[session]
    work_dir = Path(info['work_dir'])
    preview_dir = work_dir / "preview"

    # Считаем, сколько PNG уже сгенерировано
    generated_pngs = 0
    if preview_dir.exists():
        generated_pngs = len(list(preview_dir.glob("slide_*.png")))

    # Проверяем наличие PDF
    pptx_path = Path(info['original_path'])
    pdf_file = preview_dir / f"{pptx_path.stem}.pdf"
    pdf_exists = pdf_file.exists()
    pdf_size = pdf_file.stat().st_size if pdf_exists else 0

    return jsonify({
        "preview_ready": info.get('preview_ready', False),
        "preview_pages": info.get('preview_pages', 0),
        "total_slides": info.get('total_slides', 0),
        "generated_pngs": generated_pngs,
        "pdf_exists": pdf_exists,
        "pdf_size": pdf_size,
        "progress": round(generated_pngs / max(info.get('total_slides', 1), 1) * 100),
        "status": "done" if info.get('preview_ready') else ("generating" if pdf_exists else "pending")
    })

@app.route('/api/image/<session>/<int:slide_idx>', methods=['GET'])
def get_image(session, slide_idx):
    """Возвращает PNG слайда; во время генерации превью отвечает 204."""
    try:
        info = sessions_info.get(session)
        if not info:
            return jsonify({"error": "Session not found"}), 404

        if slide_idx < 0 or slide_idx >= info.get("total_slides", 0):
            return jsonify({"error": "Slide not found"}), 404

        if info.get('is_pdf'):
            img_path = IMG_DIR / session / f"slide_{slide_idx + 1}.png"
        else:
            img_path = Path(info['work_dir']) / "preview" / f"slide_{slide_idx + 1}.png"
            if not img_path.exists():
                # Совместимость с превью, созданными старыми версиями.
                img_path = IMG_DIR / session / f"slide_{slide_idx + 1}.png"

        if img_path.exists():
            return send_file(img_path, mimetype='image/png')

        if not info.get('preview_ready', False):
            return '', 204
        return jsonify({"error": "Image not found"}), 404
    except Exception as e:
        print(f"❌ Ошибка получения изображения: {e}")
        return jsonify({"error": str(e)}), 500


# В функции get_slide (API эндпоинт)
@app.route('/api/slide/<session>/<int:slide_idx>', methods=['GET'])
def get_slide(session, slide_idx):
    """
    Возвращает данные слайда: блоки с текстом, координаты, размеры.
    """
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404

    info = sessions_info[session]
    work_dir = Path(info['work_dir'])

    # Получаем список всех слайдов
    slides = slides_list(work_dir)

    if slide_idx >= len(slides):
        return jsonify({"error": f"Slide {slide_idx} not found. Total slides: {len(slides)}"}), 404

    # Получаем размер слайда в дюймах и EMU
    slide_width_inch, slide_height_inch = get_slide_size(work_dir)
    slide_width_emu = int(slide_width_inch * EMU_PER_INCH)
    slide_height_emu = int(slide_height_inch * EMU_PER_INCH)

    # Устанавливаем размер слайда для калькулятора координат
    set_slide_size_emu(slide_width_emu, slide_height_emu)

    # Очищаем кэш координат перед обработкой нового слайда
    clear_coordinate_cache()

    # Путь к XML файлу слайда
    slide_xml_path = slides[slide_idx]

    # Путь к .rels файлу (нужен для диаграмм)
    slide_rel_path = None
    rel_file_path = slide_xml_path.replace('/slides/', '/slides/_rels/') + '.rels'

    try:
        with zipfile.ZipFile(info['original_path'], 'r') as zf:
            # Проверяем существование .rels файла
            try:
                zf.getinfo(rel_file_path)
                slide_rel_path = rel_file_path
            except KeyError:
                # .rels файл не существует - это нормально для простых презентаций
                pass

            # Извлекаем блоки текста
            blocks = extract_text_blocks(slide_xml_path, zf, slide_rel_path)

    except Exception as e:
        print(f"   ❌ Ошибка извлечения блоков из слайда {slide_idx}: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Failed to extract text blocks: {str(e)}"}), 500

    # Сортируем блоки (уже отсортированы в extract_text_blocks, но на всякий случай)
    blocks.sort(key=lambda b: (b['y'], b['x']))

    # Переназначаем ID для согласованности
    for idx, block in enumerate(blocks):
        block['id'] = idx

    # Обновляем информацию о слайдах в сессии
    if 'slides_cache' not in sessions_info[session]:
        sessions_info[session]['slides_cache'] = {}

    sessions_info[session]['slides_cache'][slide_idx] = {
        'blocks': blocks,
        'width': slide_width_inch,
        'height': slide_height_inch,
        'total_slides': len(slides)
    }

    return jsonify({
        "blocks": blocks,
        "slide_width": slide_width_inch,
        "slide_height": slide_height_inch,
        "slide_num": slide_idx + 1,
        "total_slides": len(slides)
    })


@app.route('/api/generate_preview/<session>/<int:slide_idx>', methods=['POST'])
def generate_preview(session, slide_idx):
    """Генерирует превью для слайда через LibreOffice."""
    try:
        if session not in sessions_info:
            return jsonify({"error": "Session not found"}), 404

        if not LIBREOFFICE_AVAILABLE:
            return jsonify({"error": "LibreOffice not available"}), 500

        info = sessions_info[session]
        work_dir = Path(info['work_dir'])
        pptx_path = Path(info['original_path'])

        # Создаем папку preview
        preview_dir = work_dir / "preview"
        preview_dir.mkdir(parents=True, exist_ok=True)

        def generate():
            try:
                # Конвертируем PPTX в PDF
                cmd = [LIBREOFFICE_CMD, "--headless", "--convert-to", "pdf", "--outdir", str(preview_dir), str(pptx_path)]
                result = subprocess.run(cmd, capture_output=True, timeout=180)

                if result.returncode != 0:
                    print(f"LibreOffice error: {result.stderr}")
                    return

                # Конвертируем PDF в PNG
                pdf_file = preview_dir / f"{pptx_path.stem}.pdf"
                if pdf_file.exists():
                    try:
                        try:
                            import pymupdf as fitz
                        except ImportError:
                            import fitz
                        doc = fitz.open(str(pdf_file))
                        page_count = len(doc)
                        for idx, page in enumerate(doc):
                            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                            pix.save(str(preview_dir / f"slide_{idx + 1}.png"))
                        doc.close()
                        print(f"   ✅ Сгенерировано {page_count} превью для сессии {session}")
                    except ImportError:
                        print("   ⚠️ PyMuPDF не установлен")
            except Exception as e:
                print(f"   ⚠️ Ошибка генерации превью: {e}")

        # Запускаем в фоне
        threading.Thread(target=generate, daemon=True).start()

        return jsonify({"success": True, "message": "Preview generation started"})

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/apply/<session>/<int:slide_idx>', methods=['POST'])
def apply_translation(session, slide_idx):
    print(f"=== apply_translation CALLED ===")
    print(f"Session: {session}")
    print(f"Slide: {slide_idx}")

    if session not in sessions_info:
        print(f"❌ Session not found: {session}")
        return jsonify({"error": "Session not found"}), 404

    data = request.json
    print(f"Data received: {data.keys() if data else 'None'}")

    order = data.get("order", [])
    translations_list = data.get("translations", [])

    print(f"Order length: {len(order)}")
    print(f"Translations length: {len(translations_list)}")

    info = sessions_info[session]
    slides = slides_list(info['work_dir'])

    if slide_idx >= len(slides):
        print(f"❌ Slide index out of range: {slide_idx} >= {len(slides)}")
        return jsonify({"error": "Slide not found"}), 404

    xml_path = slides[slide_idx]
    print(f"XML path: {xml_path}")

    with zipfile.ZipFile(info['original_path'], 'r') as zf:
        blocks_for_slide = extract_text_blocks(xml_path, zf)

    print(f"Blocks for slide: {len(blocks_for_slide)}")

    report = apply_translation_safe(xml_path, order, translations_list, blocks_for_slide, session)
    return jsonify(report)

# ============================================================
# ФИЛЬТРЫ БЛОКОВ ДЛЯ ПРОМПТА
# ============================================================

def has_russian(text):
    """Проверяет, есть ли в тексте последовательность из 2 и более русских букв."""
    if not text:
        return False
    russian_pattern = re.compile(r'[а-яё]{2,}', re.IGNORECASE)
    return bool(russian_pattern.search(text))


def should_skip_translation(text):
    """Определяет, нужно ли переводить текст."""
    if not text or not text.strip():
        return True

    text = text.strip()

    # Только цифры, пробелы, тире, точки, запятые, проценты, $
    if re.match(r'^[\d\s\-\.\,\+\%\$]+$', text):
        return True

    # Годы (только 4 цифры)
    if re.match(r'^\d{4}$', text):
        return True

    # Комбинация букв и цифр без пробелов, длина <= 6
    if re.match(r'^[A-Za-z0-9\-\.\+\_\/]{1,6}$', text) and not re.search(r'\s', text):
        return True

    # Очень короткие (1-2 символа)
    if len(text) <= 2:
        return True

    return False


def collect_filtered_texts(session, slides_to_include, skip_russian=True):
    """
    Собирает тексты для перевода с применением фильтров.
    Возвращает список all_texts и mapping для восстановления.
    """
    info = sessions_info[session]
    all_slides = slides_list(info['work_dir'])

    all_texts = []
    mapping = []
    global_block_counter = 1

    with zipfile.ZipFile(info['original_path'], 'r') as zf:
        for slide_idx in slides_to_include:
            if slide_idx >= len(all_slides):
                continue

            blocks = extract_text_blocks(all_slides[slide_idx], zf)

            for block in blocks:
                text = block.get('text', '')
                runs = block.get('runs', [])

                # Пропуск пустых
                if not text or not text.strip():
                    continue

                # Пропуск русского текста
                if skip_russian and has_russian(text):
                    continue

                # Пропуск непереводимых
                if should_skip_translation(text):
                    continue

                has_runs_flag = len(runs) > 1
                if has_runs_flag:
                    text = ' | '.join([r.get('text', '') for r in runs])

                all_texts.append({
                    'slide_num': slide_idx + 1,
                    'block_num': global_block_counter,
                    'block_id': block['id'],
                    'text': text,
                    'has_runs': has_runs_flag,
                    'runs_count': len(runs)
                })

                mapping.append({
                    'slide_num': slide_idx + 1,
                    'block_num': global_block_counter,
                    'block_id': block['id']
                })

                global_block_counter += 1

    return all_texts, mapping


def collect_all_texts_with_meta(session):
    """Совместимый сбор всех переводимых блоков презентации."""
    info = sessions_info[session]
    all_texts, _ = collect_filtered_texts(
        session, range(len(slides_list(info["work_dir"]))), skip_russian=True
    )
    return all_texts


def _text_batches(texts, max_chars=PROMPT_BATCH_CHARS):
    """Разбивает запросы по объёму исходного текста, не разрывая блоки."""
    batches, current, size = [], [], 0
    for item in texts:
        item_size = len(item.get("text", "")) + 40
        if current and size + item_size > max_chars:
            batches.append(current)
            current, size = [], 0
        current.append(item)
        size += item_size
    if current:
        batches.append(current)
    return batches


def _perform_translation(session, slide_indexes, data):
    """Единый цикл перевода для всех API: сбор, запросы, проверка, сохранение."""
    info = sessions_info[session]
    data = data or {}
    api_key = data.get("api_key") or USER_SETTINGS.get("api_key", "")
    if not api_key:
        raise ValueError("API key required")

    slide_count = len(slides_list(info["work_dir"]))
    slides = sorted({int(index) for index in slide_indexes})
    if not slides or any(index < 0 or index >= slide_count for index in slides):
        raise ValueError("Invalid slides selection")

    all_texts, mapping = collect_filtered_texts(session, slides, skip_russian=True)
    if not all_texts:
        raise ValueError("No blocks to translate")

    enabled_variant_ids = data.get("enabled_variants")
    translations_matrix = []
    active_variants = None
    batches = _text_batches(all_texts)
    for batch_number, batch in enumerate(batches, 1):
        print(f"   🤖 Запрос {batch_number}/{len(batches)}: {len(batch)} блоков")
        prompt, batch_variants = _build_all_slides_prompt(
            batch, data.get("target_lang", "ru"), enabled_variant_ids
        )
        if active_variants is None:
            active_variants = batch_variants
        elif [v["id"] for v in active_variants] != [v["id"] for v in batch_variants]:
            raise RuntimeError("Active variants changed between requests")

        response = _call_deepseek_single(
            api_key,
            data.get("model") or USER_SETTINGS.get("model", "deepseek-reasoner"),
            prompt,
            data.get("proxy_url") or USER_SETTINGS.get("proxy_url", ""),
            session,
        )
        translations_matrix.extend(_parse_translations_matrix(
            response, len(batch), len(active_variants), strict=True
        ))

    active_variants = active_variants or []
    info.setdefault("translations", {})
    translated_blocks = 0
    for item, row in zip(all_texts, translations_matrix):
        slide_idx = item["slide_num"] - 1
        block_translations = {}
        for variant, value in zip(active_variants, row):
            if value:
                block_translations[f"{variant['id']}_translation"] = value
                block_translations[f"{variant['id']}_source"] = "api"
        if block_translations:
            info["translations"].setdefault(slide_idx, {})[str(item["block_id"])] = block_translations
            translated_blocks += 1

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    mapping_file = OUT_DIR / f"mapping_{session[:8]}_{timestamp}.json"
    mapping_file.write_text(json.dumps({
        "mapping": mapping,
        "active_variants": active_variants,
        "blocks": all_texts,
        "translations_matrix": translations_matrix,
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    translations_for_json = {
        str(slide_idx): slide_trans
        for slide_idx, slide_trans in info["translations"].items()
    }
    return {
        "success": True,
        "total_slides": len(slides),
        "translated_blocks": translated_blocks,
        "active_variants": active_variants,
        "mapping_file": mapping_file.name,
        "translations": translations_for_json,
    }


# ============================================================
# ВЫЗОВ DEEPSEEK API
# ============================================================

def _call_deepseek_single(api_key, model, prompt, proxy_url, session_id=None):
    """Выполняет один запрос к DeepSeek с проверкой TLS по умолчанию."""
    print("   ⏳ Отправка запроса к DeepSeek API...")
    print(f"   📏 Размер промпта: {len(prompt)} символов (~{len(prompt)//4} токенов)")

    proxies = {"http": proxy_url, "https": proxy_url} if proxy_url else None
    start_time = time.time()
    response = requests.post(
        f"{DEEPSEEK_BASE_URL}/chat/completions",
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json={
            "model": model,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.3,
            "stream": False,
        },
        timeout=(20, 300),
        proxies=proxies,
        verify=DEEPSEEK_VERIFY_TLS,
    )
    print(f"   ⏱️ Запрос выполнен за {time.time() - start_time:.1f} секунд")
    print(f"   📡 Статус ответа: {response.status_code}")

    if response.status_code != 200:
        error_text = response.text[:500]
        print(f"   ❌ API Error {response.status_code}: {error_text}")
        raise RuntimeError(f"DeepSeek API error {response.status_code}: {error_text}")

    data = response.json()
    try:
        content = data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise ValueError("DeepSeek returned an unexpected response") from exc

    usage = data.get("usage", {})
    if usage:
        print(f"   📊 Токены: {usage.get('total_tokens', 0)}")

    if session_id:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        response_file = OUT_DIR / f"api_response_{session_id[:8]}_{timestamp}.txt"
        response_file.write_text(content, encoding="utf-8")
        print(f"   💾 Ответ сохранен: {response_file.name}")

    return content


# ============================================================
# ПАРСИНГ ОТВЕТА (С РЕМОНТОМ JSON)
# ============================================================

def _repair_truncated_json(json_text):
    print(f"   🔧 Попытка восстановления оборванного JSON...")

    lines = json_text.split('\n')
    for i in range(len(lines) - 1, -1, -1):
        line = lines[i].strip()
        if line.endswith(('",', '"]', ']', '"')):
            lines = lines[:i + 1]
            break

    json_text = '\n'.join(lines)

    in_string = False
    last_complete_pos = 0
    for i, char in enumerate(json_text):
        if char == '"' and (i == 0 or json_text[i - 1] != '\\'):
            in_string = not in_string
            if not in_string:
                last_complete_pos = i + 1

    if in_string and last_complete_pos > 0:
        json_text = json_text[:last_complete_pos]
        print(f"   🔧 Обрезана незакрытая строка")

    json_text = json_text.rstrip()
    if json_text.endswith(','):
        json_text = json_text[:-1]

    open_curly = json_text.count('{')
    close_curly = json_text.count('}')
    open_square = json_text.count('[')
    close_square = json_text.count(']')

    json_text += ']' * (open_square - close_square)
    json_text += '}' * (open_curly - close_curly)

    return json_text


def _parse_translations_matrix(response_text, expected_blocks, expected_variants, strict=False):
    import re

    print(f"   🔍 Поиск JSON в ответе...")

    clean_text = response_text
    if '```json' in clean_text:
        clean_text = clean_text.split('```json', 1)[1]
    elif '```' in clean_text:
        clean_text = clean_text.split('```', 1)[1]
    if '```' in clean_text:
        clean_text = clean_text.rsplit('```', 1)[0]
    clean_text = clean_text.strip()

    matrix = []

    try:
        data = json.loads(clean_text)
        matrix = data.get('translations', [])
        print(f"   ✅ JSON распарсен, блоков: {len(matrix)}")
    except json.JSONDecodeError as e:
        print(f"   ⚠️ Ошибка JSON: {e}")
        repaired = _repair_truncated_json(clean_text)
        try:
            data = json.loads(repaired)
            matrix = data.get('translations', [])
            print(f"   ✅ JSON отремонтирован, блоков: {len(matrix)}")
        except json.JSONDecodeError:
            print(f"   ❌ Не удалось восстановить JSON")
            matrix = []

    if strict and len(matrix) != expected_blocks:
        raise ValueError(
            f"Translation response contains {len(matrix)} blocks; expected {expected_blocks}"
        )

    normalized = []
    for i, row in enumerate(matrix[:expected_blocks]):
        if not isinstance(row, list):
            if strict:
                raise ValueError(f"Translation block {i + 1} is not an array")
            row = [str(row)] if row else []
        if strict and len(row) != expected_variants:
            raise ValueError(
                f"Translation block {i + 1} contains {len(row)} variants; "
                f"expected {expected_variants}"
            )
        row = ["" if value is None else str(value) for value in row]
        while len(row) < expected_variants:
            row.append("")
        normalized.append(row[:expected_variants])

    while len(normalized) < expected_blocks:
        if strict:
            raise ValueError("Translation response is incomplete")
        normalized.append([""] * expected_variants)
    return normalized


# ============================================================
# СОХРАНЕНИЕ MAPPING И ЗАГРУЗКА ИЗ ФАЙЛА
# ============================================================

@app.route('/api/translate_with_filters/<session>', methods=['POST'])
def translate_with_filters(session):
    """Перевод выбранных слайдов; активные стили едины для запроса и ответа."""
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404
    try:
        data = request.get_json(silent=True) or {}
        return jsonify(_perform_translation(session, data.get("slides", []), data))
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        print(f"❌ Ошибка перевода: {exc}")
        traceback.print_exc()
        return jsonify({"error": str(exc)}), 500


@app.route('/api/load_response_from_file', methods=['POST'])
def load_response_from_file():
    """Импортирует ответ и mapping через multipart upload, без путей клиента."""
    session = request.form.get("session", "")
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404
    response_upload = request.files.get("response_file")
    mapping_upload = request.files.get("mapping_file")
    if not response_upload or not mapping_upload:
        return jsonify({"error": "Response and mapping files are required"}), 400

    try:
        response_text = response_upload.read().decode("utf-8-sig")
        mapping_data = json.loads(mapping_upload.read().decode("utf-8-sig"))
        mapping = mapping_data["mapping"]
        active_variants = mapping_data["active_variants"]
        if not isinstance(mapping, list) or not isinstance(active_variants, list):
            raise ValueError("Invalid mapping structure")
        if not active_variants or any(not isinstance(v, dict) or not v.get("id") for v in active_variants):
            raise ValueError("Invalid variants in mapping")
        for item in mapping:
            if not isinstance(item, dict) or "slide_num" not in item or "block_id" not in item:
                raise ValueError("Invalid block in mapping")
            slide_idx = int(item["slide_num"]) - 1
            if slide_idx < 0 or slide_idx >= sessions_info[session].get("total_slides", 0):
                raise ValueError("Mapping refers to a missing slide")

        translations_matrix = _parse_translations_matrix(
            response_text, len(mapping), len(active_variants), strict=True
        )
        info = sessions_info[session]
        info.setdefault("translations", {})
        saved = 0
        for item, row in zip(mapping, translations_matrix):
            slide_idx = int(item["slide_num"]) - 1
            block_trans = {}
            for variant, value in zip(active_variants, row):
                if value:
                    block_trans[f"{variant['id']}_translation"] = value
                    block_trans[f"{variant['id']}_source"] = "file"
            if block_trans:
                info["translations"].setdefault(slide_idx, {})[str(item["block_id"])] = block_trans
                saved += 1

        return jsonify({
            "success": True,
            "saved_blocks": saved,
            "translations": {
                str(slide_idx): slide_trans
                for slide_idx, slide_trans in info["translations"].items()
            },
        })
    except (ValueError, KeyError, UnicodeError, json.JSONDecodeError) as exc:
        return jsonify({"error": f"Invalid import: {exc}"}), 400
    except Exception as exc:
        print(f"❌ Ошибка загрузки из файла: {exc}")
        traceback.print_exc()
        return jsonify({"error": str(exc)}), 500


@app.route('/api/translate/all_single/<session>', methods=['POST'])
def translate_all_single(session):
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404
    try:
        data = request.get_json(silent=True) or {}
        slide_indexes = range(sessions_info[session].get("total_slides", 0))
        return jsonify(_perform_translation(session, slide_indexes, data))
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        print(f"❌ Ошибка перевода презентации: {exc}")
        traceback.print_exc()
        return jsonify({"error": str(exc)}), 500


@app.route('/api/deepseek/status', methods=['GET'])
def deepseek_status():
    """Проверяет доступность DeepSeek API."""
    try:
        api_key = USER_SETTINGS.get('api_key', '')
        if not api_key:
            return jsonify({"status": "no_key", "message": "API ключ не задан"})

        # Простой тестовый запрос
        response = requests.post(
            f"{DEEPSEEK_BASE_URL}/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={
                "model": "deepseek-reasoner",
                "messages": [{"role": "user", "content": "OK"}],
                "max_tokens": 1
            },
            timeout=10,
            verify=DEEPSEEK_VERIFY_TLS
        )

        if response.status_code == 200:
            return jsonify({"status": "ok", "message": "DeepSeek API доступен"})
        else:
            return jsonify({"status": "error", "message": f"HTTP {response.status_code}", "detail": response.text[:200]})

    except requests.exceptions.Timeout:
        return jsonify({"status": "timeout", "message": "Таймаут соединения"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

@app.route('/api/variants/list', methods=['GET'])
def api_get_variants_list():
    """Возвращает список вариантов из TXT файла."""
    try:
        variants = get_variants_list()
        return jsonify({"success": True, "variants": variants})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/translate/slide_single/<session>/<int:slide_idx>', methods=['POST'])
def translate_slide_single(session, slide_idx):
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404
    try:
        data = request.get_json(silent=True) or {}
        return jsonify(_perform_translation(session, [slide_idx], data))
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        print(f"❌ Ошибка перевода слайда: {exc}")
        traceback.print_exc()
        return jsonify({"error": str(exc)}), 500


@app.route('/api/apply_loaded_response/<session>', methods=['POST'])
def apply_loaded_response(session):
    """Применяет загруженный из файла ответ к сессии."""
    try:
        if session not in sessions_info:
            return jsonify({"error": "Session not found"}), 404

        full_response = request.get_json()

        blocks = full_response.get('blocks', [])
        translations_matrix = full_response.get('translations_matrix', [])
        active_variants = full_response.get('active_variants', [])

        info = sessions_info[session]

        # Сохраняем переводы
        if 'translations' not in info:
            info['translations'] = {}

        for i, block in enumerate(blocks):
            slide_idx = block['slide_num'] - 1
            block_id = str(block['block_id'])

            if slide_idx not in info['translations']:
                info['translations'][slide_idx] = {}

            block_translations = {}
            if i < len(translations_matrix):
                for v_idx, variant in enumerate(active_variants):
                    if v_idx < len(translations_matrix[i]):
                        block_translations[f"{variant['id']}_translation"] = translations_matrix[i][v_idx]
                        block_translations[f"{variant['id']}_source"] = 'file'

            info['translations'][slide_idx][block_id] = block_translations

        # Приводим ключи к строкам
        translations_for_json = {}
        for slide_idx, slide_trans in info['translations'].items():
            translations_for_json[str(slide_idx)] = slide_trans

        return jsonify({
            "success": True,
            "translations": translations_for_json
        })

    except Exception as e:
        print(f"❌ Ошибка применения ответа: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/export_prompt/<session>', methods=['GET'])
def export_prompt(session):
    """Экспорт полного промпта для проверки."""
    try:
        if session not in sessions_info:
            return jsonify({"error": "Session not found"}), 404

        target_lang = request.args.get('lang', 'ru')

        all_texts_with_meta = collect_all_texts_with_meta(session)
        if not all_texts_with_meta:
            return jsonify({"error": "No text blocks found"}), 400

        # Получаем включенные варианты из настроек
        enabled_variant_ids = USER_SETTINGS.get('enabled_variants', None)

        # _build_all_slides_prompt возвращает КОРТЕЖ (prompt, active_variants)
        prompt, active_variants = _build_all_slides_prompt(
            all_texts_with_meta,
            target_lang,
            enabled_variant_ids
        )

        # Сохраняем во временный файл
        prompt_file = OUT_DIR / f"prompt_{session}.txt"
        with open(prompt_file, 'w', encoding='utf-8') as f:
            f.write(prompt)

        return send_file(
            prompt_file,
            as_attachment=True,
            download_name=f"deepseek_prompt_slide_all.txt",
            mimetype='text/plain'
        )

    except Exception as e:
        print(f"❌ Ошибка экспорта промпта: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route('/api/glossary/active/<session>/<int:slide_idx>', methods=['GET'])
def get_active_glossary(session, slide_idx):
    if session not in sessions_info:
        return jsonify({"error": "Session not found"}), 404

    info = sessions_info[session]
    slides = slides_list(info['work_dir'])
    if slide_idx >= len(slides):
        return jsonify({"error": "Slide not found"}), 404

    with zipfile.ZipFile(info['original_path'], 'r') as zf:
        blocks = extract_text_blocks(slides[slide_idx], zf)

    all_text = ' '.join([b.get('text', '') for b in blocks]).lower()
    glossary_terms = load_glossary()

    active_terms = []
    for term in glossary_terms:
        if term['source'].lower() in all_text:
            idx = all_text.find(term['source'].lower())
            start = max(0, idx - 30)
            end = min(len(all_text), idx + len(term['source']) + 30)
            context = all_text[start:end]
            active_terms.append({
                'source': term['source'], 'target': term['target'],
                'found_in': ('...' if start > 0 else '') + context + ('...' if end < len(all_text) else '')
            })

    return jsonify({"success": True, "slide_num": slide_idx + 1,
                    "total_terms_in_glossary": len(glossary_terms),
                    "active_terms_count": len(active_terms), "active_terms": active_terms})


@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({"ok": True, "libreoffice_available": LIBREOFFICE_AVAILABLE,
                    "api_key_configured": bool(USER_SETTINGS.get('api_key'))})


# ========== ЗАПУСК ==========
if __name__ == '__main__':
    print("\n" + "=" * 50)
    print("🚀 PPTX Translation Tool v5.0")
    print("=" * 50)
    print(f"LibreOffice: {'✅' if LIBREOFFICE_AVAILABLE else '❌'}")
    print(f"Рабочая директория: {WORKSPACE_DIR}")
    print("📍 http://localhost:5001")
    print("=" * 50 + "\n")
    app.run(host=APP_HOST, port=APP_PORT, debug=os.getenv("FLASK_DEBUG", "false").lower() == "true")