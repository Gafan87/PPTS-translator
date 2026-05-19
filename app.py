from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from pathlib import Path
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
from pathlib import Path
from lxml import etree
import copy
from dotenv import load_dotenv
import requests
from datetime import datetime
import traceback 

# ========== ОТКЛЮЧАЕМ ПРЕДУПРЕЖДЕНИЯ SSL ==========
import ssl
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Отключаем логи запросов Flask
import logging
logging.getLogger('werkzeug').setLevel(logging.ERROR)  # ← Только ошибки
# ИЛИ logging.getLogger('werkzeug').disabled = True     # ← Полностью отключить

# ========== ИНИЦИАЛИЗАЦИЯ ==========
app = Flask(__name__)
CORS(app)

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

APP_HOST = os.getenv("APP_HOST", "127.0.0.1")
APP_PORT = int(os.getenv("APP_PORT", "5001"))
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1")

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
            if created and (now - created).seconds > 86400:  # 24 часа
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

def save_session(session_id, custom_name=None):
    """Сохраняет текущую сессию в папку saved_sessions."""
    if session_id not in sessions_info:
        return {"success": False, "error": "Session not found"}
    
    info = sessions_info[session_id]
    original_name = info.get('original_name', 'presentation')
    
    # Формируем имя папки: название_презентации_дата
    date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder_name = f"{original_name}_{date_str}"
    if custom_name:
        folder_name = custom_name
    
    saved_dir = SESSIONS_DIR / folder_name
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
    
    # ========== СОХРАНЯЕМ ПЕРЕВОДЫ В TXT ==========
    translations_file = saved_dir / "translations.txt"
    if 'translations' in info and info['translations']:
        export_translations_to_txt(info['translations'], info.get('total_slides', 0), translations_file)
        print(f"   💾 Переводы сохранены в {translations_file}")
    # =============================================
    
    # Сохраняем метаданные
    meta = {
        "session_id": session_id,
        "original_name": original_name,
        "original_path": str(dest_original),
        "work_dir": str(dest_work),
        "total_slides": info.get('total_slides', 0),
        "saved_at": datetime.now().isoformat(),
        "custom_name": folder_name,
        "has_translations": 'translations' in info,
        "translations_file": str(translations_file) if 'translations' in info else None
    }
    
    # Сохраняем маркеры и области
    if 'markers' in info:
        meta['markers'] = info['markers']
    if 'areas' in info:
        meta['areas'] = info['areas']
    
    with open(saved_dir / "session_info.json", 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    
    return {"success": True, "folder": folder_name, "path": str(saved_dir)}

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
            slide_trans = translations.get(str(slide_idx), {})
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
            work_exists = Path(meta.get('work_dir', '')).exists()
            original_exists = Path(meta.get('original_path', '')).exists()
            
            sessions.append({
                "folder": folder.name,
                "original_name": meta.get('original_name', ''),
                "saved_at": meta.get('saved_at', ''),
                "total_slides": meta.get('total_slides', 0),
                "has_translations": 'translations' in meta,
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
    saved_dir = SESSIONS_DIR / folder_name
    info_file = saved_dir / "session_info.json"
    
    if not info_file.exists():
        return {"success": False, "error": "Session info not found"}
    
    with open(info_file, 'r', encoding='utf-8') as f:
        meta = json.load(f)
    
    session_id = new_session_id or str(uuid.uuid4())
    
    # Копируем work_dir
    work_dir = WORK_DIR / session_id
    saved_work = Path(meta['work_dir'])
    if saved_work.exists():
        if work_dir.exists():
            shutil.rmtree(work_dir)
        shutil.copytree(saved_work, work_dir)
    
    # Копируем оригинальный файл
    original_path = UPLOAD_DIR / f"{session_id}.pptx"
    saved_original = Path(meta['original_path'])
    if saved_original.exists():
        shutil.copy2(saved_original, original_path)
    
    # ========== ЗАГРУЖАЕМ ПЕРЕВОДЫ ИЗ TXT ==========
    translations = {}
    translations_file = saved_dir / "translations.txt"
    if translations_file.exists():
        translations = import_translations_from_txt(translations_file)
        print(f"   📂 Переводы загружены из {translations_file}")
    # =============================================
    
    # Восстанавливаем сессию
    sessions_info[session_id] = {
        'original_name': meta.get('original_name', ''),
        'original_path': str(original_path),
        'work_dir': str(work_dir),
        'total_slides': meta.get('total_slides', 0),
        'is_pdf': False,
        'translations': translations
    }
    
    if 'markers' in meta:
        sessions_info[session_id]['markers'] = meta['markers']
    if 'areas' in meta:
        sessions_info[session_id]['areas'] = meta['areas']
    
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
        saved_dir = SESSIONS_DIR / folder_name
        
        # Удаляем старую папку если существует
        if saved_dir.exists():
            shutil.rmtree(saved_dir)
            print(f"   🔄 Перезапись сессии: {folder_name}")
    else:
        # Обычное сохранение (новая сессия)
        if custom_name:
            folder_name = custom_name
        else:
            folder_name = f"{original_name}_{date_str}"
    
    saved_dir = SESSIONS_DIR / folder_name
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
    
    # Сохраняем переводы в TXT
    translations_file = saved_dir / "translations.txt"
    if 'translations' in info and info['translations']:
        export_translations_to_txt(info['translations'], info.get('total_slides', 0), translations_file)
        print(f"   💾 Переводы сохранены в {translations_file}")
    
    # Сохраняем метаданные
    meta = {
        "session_id": session_id,
        "original_name": original_name,
        "original_path": str(dest_original),
        "work_dir": str(dest_work),
        "total_slides": info.get('total_slides', 0),
        "saved_at": datetime.now().isoformat(),
        "custom_name": folder_name,
        "has_translations": 'translations' in info,
        "translations_file": str(translations_file) if 'translations' in info else None,
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
        saved_dir = SESSIONS_DIR / folder_name
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
    with zipfile.ZipFile(path, 'r') as z:
        z.extractall(dest)

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

# Глобальный кэш для калькуляторов групп
_coordinate_calculator_cache = {}
_slide_width_emu = 9144000   # 10 дюймов по умолчанию
_slide_height_emu = 6858000  # 7.5 дюймов по умолчанию

def set_slide_size_emu(width_emu, height_emu):
    """Устанавливает размер слайда в EMU для проверки координат."""
    global _slide_width_emu, _slide_height_emu
    _slide_width_emu = width_emu
    _slide_height_emu = height_emu

def clear_coordinate_cache():
    """Очищает кэш калькуляторов координат (для каждого нового слайда)."""
    global _coordinate_calculator_cache
    _coordinate_calculator_cache.clear()

def _get_local_coordinates(element):
    """Возвращает локальные координаты элемента (без учета групп)."""
    xfrm = element.find('.//p:xfrm', namespaces=NAMESPACES)
    if xfrm is None:
        xfrm = element.find('.//a:xfrm', namespaces=NAMESPACES)
    
    if xfrm is not None:
        off = xfrm.find('.//a:off', namespaces=NAMESPACES)
        if off is not None:
            return int(off.get('x', 0)), int(off.get('y', 0))
    return 0, 0

def _get_element_size(element):
    """Возвращает размер элемента (ширина, высота) в EMU."""
    xfrm = element.find('.//p:xfrm', namespaces=NAMESPACES)
    if xfrm is None:
        xfrm = element.find('.//a:xfrm', namespaces=NAMESPACES)
    
    if xfrm is not None:
        ext = xfrm.find('.//a:ext', namespaces=NAMESPACES)
        if ext is not None:
            return int(ext.get('cx', 0)), int(ext.get('cy', 0))
    return 0, 0

def _get_parent_offsets(element):
    """
    Собирает смещения всех родительских групп.
    Возвращает список от самой внешней к самой внутренней.
    """
    offsets = []
    current = element.getparent()
    
    while current is not None:
        if current.tag.endswith('grpSp'):
            xfrm = current.find('.//a:xfrm', namespaces=NAMESPACES)
            if xfrm is None:
                xfrm = current.find('.//p:xfrm', namespaces=NAMESPACES)
            
            if xfrm is not None:
                off = xfrm.find('.//a:off', namespaces=NAMESPACES)
                if off is not None:
                    offsets.insert(0, {
                        'x': int(off.get('x', 0)),
                        'y': int(off.get('y', 0))
                    })
        current = current.getparent()
    
    return offsets

def _find_best_depth_for_group(group_element):
    """
    Определяет оптимальную глубину суммирования для всех блоков в группе.
    Возвращает количество родительских групп, чьи смещения нужно суммировать.
    """
    # Находим все текстовые блоки в группе
    text_blocks = group_element.xpath('.//p:sp | .//p:graphicFrame', namespaces=NAMESPACES)
    
    if not text_blocks:
        return None
    
    # Собираем смещения родительских групп
    parent_offsets = _get_parent_offsets(group_element)
    
    if not parent_offsets:
        return 0
    
    best_depth = None
    best_score = -1
    
    # Пробуем разные глубины суммирования
    for depth in range(0, len(parent_offsets) + 1):
        all_blocks_valid = True
        total_score = 0
        
        for block in text_blocks:
            local_x, local_y = _get_local_coordinates(block)
            block_w, block_h = _get_element_size(block)
            
            # Суммируем с выбранной глубиной
            sum_x = local_x
            sum_y = local_y
            for i in range(depth):
                sum_x += parent_offsets[i]['x']
                sum_y += parent_offsets[i]['y']
            
            # Проверяем попадание в слайд
            x_inside = 0 <= sum_x <= _slide_width_emu
            y_inside = 0 <= sum_y <= _slide_height_emu
            
            # Блок может немного выходить за границы (до 10%)
            margin_x = _slide_width_emu * 0.1
            margin_y = _slide_height_emu * 0.1
            x_near = -margin_x <= sum_x <= _slide_width_emu + margin_x
            y_near = -margin_y <= sum_y <= _slide_height_emu + margin_y
            
            if x_inside and y_inside:
                # Блок внутри слайда - отлично
                center_x = _slide_width_emu // 2
                center_y = _slide_height_emu // 2
                distance = abs(sum_x - center_x) + abs(sum_y - center_y)
                total_score += 10000 - distance
            elif x_near and y_near:
                # Блок у границы - допустимо
                total_score += 1000
            else:
                # Блок за пределами - этот вариант не подходит
                all_blocks_valid = False
                break
        
        if all_blocks_valid and total_score > best_score:
            best_score = total_score
            best_depth = depth
    
    return best_depth if best_depth is not None else len(parent_offsets)

def get_absolute_coordinates(element):
    """
    Универсальное вычисление координат для ВСЕХ сценариев.
    """
    
    # 1. Получаем локальные координаты блока (если есть)
    local_x, local_y = _get_local_coordinates(element)
    local_w, local_h = _get_element_size(element)
    
    # 2. Собираем ВСЕ родительские группы с их параметрами
    groups = []
    current = element.getparent()
    
    while current is not None:
        if current.tag.endswith('grpSp'):
            group_x, group_y = _get_group_offset(current)
            group_w, group_h = _get_group_size(current)
            ch_x, ch_y = _get_group_choff(current)
            
            groups.insert(0, {
                'element': current,
                'offset_x': group_x,
                'offset_y': group_y,
                'width': group_w,
                'height': group_h,
                'choff_x': ch_x,
                'choff_y': ch_y
            })
        current = current.getparent()
    
    # 3. Если нет групп - возвращаем локальные координаты
    if not groups:
        return local_x, local_y
    
    # 4. Генерируем ВСЕ возможные варианты координат
    candidates = []
    
    # Вариант 1: локальные координаты (если есть)
    if local_x != 0 or local_y != 0:
        candidates.append({
            'x': local_x,
            'y': local_y,
            'type': 'локальные',
            'score': 0
        })
    
    # Вариант 2: координаты группы (offset и chOff)
    for i, group in enumerate(groups):
        candidates.append({
            'x': group['offset_x'],
            'y': group['offset_y'],
            'type': f'группа {i+1} (offset)',
            'depth': i + 1,
            'score': 0
        })
        
        if group['choff_x'] != 0 or group['choff_y'] != 0:
            candidates.append({
                'x': group['choff_x'],
                'y': group['choff_y'],
                'type': f'группа {i+1} (chOff)',
                'depth': i + 1,
                'score': 0
            })
    
    # Вариант 3: суммы с разными уровнями групп
    for depth in range(1, len(groups) + 1):
        sum_x = local_x
        sum_y = local_y
        for i in range(depth):
            sum_x += groups[i]['offset_x']
            sum_y += groups[i]['offset_y']
        candidates.append({
            'x': sum_x,
            'y': sum_y,
            'type': f'сумма {depth} групп',
            'depth': depth,
            'score': 0
        })
    
    # Вариант 4: сумма с chOff
    for depth in range(1, len(groups) + 1):
        sum_x = local_x
        sum_y = local_y
        for i in range(depth):
            sum_x += groups[i]['choff_x']
            sum_y += groups[i]['choff_y']
        candidates.append({
            'x': sum_x,
            'y': sum_y,
            'type': f'сумма chOff {depth} групп',
            'depth': depth,
            'score': 0
        })
    
    # 5. Оцениваем каждый вариант
    slide_w, slide_h = _get_slide_size_emu()
    
    for cand in candidates:
        score = 0
        
        # Критерий 1: попадание в слайд
        if 0 <= cand['x'] <= slide_w and 0 <= cand['y'] <= slide_h:
            score += 100
        elif -slide_w*0.1 <= cand['x'] <= slide_w*1.1 and -slide_h*0.1 <= cand['y'] <= slide_h*1.1:
            score += 50
        
        # Критерий 2: проверка попадания в родительские группы
        for i, group in enumerate(groups):
            in_group = (group['offset_x'] <= cand['x'] <= group['offset_x'] + group['width'] and
                       group['offset_y'] <= cand['y'] <= group['offset_y'] + group['height'])
            if in_group:
                score += 30 * (i + 1)
        
        # Критерий 3: для блоков без своих координат - предпочитаем chOff
        if local_x == 0 and local_y == 0 and local_w == 0 and local_h == 0:
            if 'chOff' in cand['type']:
                score += 200
            elif 'offset' in cand['type']:
                score += 50
        
        # ========== НОВЫЙ КРИТЕРИЙ (только когда есть локальные координаты) ==========
        if (local_x != 0 or local_y != 0) and groups:
            # Берем первую (самую внешнюю) группу
            outer_group = groups[0]
            
            # Проверяем, находятся ли ЛОКАЛЬНЫЕ координаты ВНЕ группы
            local_outside = (local_x < outer_group['offset_x'] or 
                            local_x > outer_group['offset_x'] + outer_group['width'] or
                            local_y < outer_group['offset_y'] or 
                            local_y > outer_group['offset_y'] + outer_group['height'])
            
            if local_outside:
                # Если локальные координаты вне группы - они уже глобальные
                if cand['type'] == 'локальные':
                    score += 150  # добавляем бонус, но не перебиваем другие критерии
        
        # Критерий 4: координаты не должны быть слишком большими
        if abs(cand['x']) > slide_w * 2 or abs(cand['y']) > slide_h * 2:
            score -= 100
        
        cand['score'] = score
    
    # 6. Выбираем лучший вариант
    best = max(candidates, key=lambda c: c['score'])
    
    print(f"   🎯 {best['type']}: ({best['x']/EMU_PER_INCH:.2f}, {best['y']/EMU_PER_INCH:.2f}) дюймов (score={best['score']})")
    
    return best['x'], best['y']

def _get_group_choff(group_element):
    """Возвращает смещение дочерних элементов (chOff) группы в EMU."""
    xfrm = group_element.find('.//a:xfrm', namespaces=NAMESPACES)
    if xfrm is None:
        xfrm = group_element.find('.//p:xfrm', namespaces=NAMESPACES)
    
    if xfrm is not None:
        chOff = xfrm.find('.//a:chOff', namespaces=NAMESPACES)
        if chOff is not None:
            return int(chOff.get('x', 0)), int(chOff.get('y', 0))
    return 0, 0

def _get_group_offset(group_element):
    """Возвращает смещение группы в EMU."""
    xfrm = group_element.find('.//a:xfrm', namespaces=NAMESPACES)
    if xfrm is None:
        xfrm = group_element.find('.//p:xfrm', namespaces=NAMESPACES)
    if xfrm is None:
        xfrm = group_element.find('.//p:grpSpPr//a:xfrm', namespaces=NAMESPACES)
    
    if xfrm is not None:
        off = xfrm.find('.//a:off', namespaces=NAMESPACES)
        if off is not None:
            return int(off.get('x', 0)), int(off.get('y', 0))
    return 0, 0

def _get_group_size(group_element):
    """Возвращает размер группы в EMU."""
    xfrm = group_element.find('.//a:xfrm', namespaces=NAMESPACES)
    if xfrm is None:
        xfrm = group_element.find('.//p:xfrm', namespaces=NAMESPACES)
    if xfrm is None:
        xfrm = group_element.find('.//p:grpSpPr//a:xfrm', namespaces=NAMESPACES)
    
    if xfrm is not None:
        ext = xfrm.find('.//a:ext', namespaces=NAMESPACES)
        if ext is not None:
            return int(ext.get('cx', 0)), int(ext.get('cy', 0))
    return 0, 0

def _get_element_size(element):
    """Возвращает размер элемента в EMU."""
    xfrm = element.find('.//p:xfrm', namespaces=NAMESPACES)
    if xfrm is None:
        xfrm = element.find('.//a:xfrm', namespaces=NAMESPACES)
    
    if xfrm is not None:
        ext = xfrm.find('.//a:ext', namespaces=NAMESPACES)
        if ext is not None:
            return int(ext.get('cx', 0)), int(ext.get('cy', 0))
    return 0, 0

# Глобальные переменные для размера слайда
_slide_width_emu = 9144000   # 10 дюймов по умолчанию
_slide_height_emu = 6858000  # 7.5 дюймов по умолчанию

def _get_slide_size_emu():
    """Возвращает размер слайда в EMU."""
    return _slide_width_emu, _slide_height_emu

def set_slide_size_emu(width_emu, height_emu):
    """Устанавливает размер слайда в EMU."""
    global _slide_width_emu, _slide_height_emu
    _slide_width_emu = width_emu
    _slide_height_emu = height_emu
    
# ============================================================
# ДИАГРАММЫ
# ============================================================

def extract_chart_text_by_path(pptx_zip, chart_path, container):
    """Извлекает текст из диаграммы по точному пути."""
    try:
        with pptx_zip.open(chart_path) as f:
            chart_tree = etree.fromstring(f.read())
        
        chart_ns = {
            'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
        }
        
        runs = []
        
        # Заголовок
        for t_elem in chart_tree.xpath('//c:title//a:t', namespaces=chart_ns):
            if t_elem.text and t_elem.text.strip():
                runs.append({'text': t_elem.text.strip(), 'formatting': {}})
        
        # Категории (оси X)
        for v_elem in chart_tree.xpath('//c:cat//c:strRef//c:strCache//c:pt//c:v', namespaces=chart_ns):
            if v_elem.text and v_elem.text.strip():
                runs.append({'text': v_elem.text.strip(), 'formatting': {}})
        
        # Значения (оси Y)
        for v_elem in chart_tree.xpath('//c:val//c:numRef//c:numCache//c:pt//c:v', namespaces=chart_ns):
            if v_elem.text and v_elem.text.strip():
                runs.append({'text': v_elem.text.strip(), 'formatting': {}})
        
        # Легенда
        for t_elem in chart_tree.xpath('//c:legend//c:txPr//a:p//a:r//a:t', namespaces=chart_ns):
            if t_elem.text and t_elem.text.strip():
                runs.append({'text': t_elem.text.strip(), 'formatting': {}})
        
        if not runs:
            return None
        
        # Координаты и размер
        abs_x, abs_y = get_absolute_coordinates(container)
        w, h = _get_element_size(container)
        
        return {
            "runs": runs,
            "text": ' | '.join([r['text'] for r in runs]),
            "x": abs_x / EMU_PER_INCH,
            "y": abs_y / EMU_PER_INCH,
            "w": w / EMU_PER_INCH,
            "h": h / EMU_PER_INCH,
            "chart_file": chart_path,
            "type": "chart"
        }
    except Exception as e:
        print(f"   ⚠️ Ошибка диаграммы {chart_path}: {e}")
        return None

def extract_chart_text(pptx_zip, chart_rId, container, slide_rel_path):
    """Извлекает текст из диаграммы по rId через .rels файл."""
    if not slide_rel_path:
        return None
    
    try:
        with pptx_zip.open(slide_rel_path) as f:
            rel_tree = etree.fromstring(f.read())
        
        rel_ns = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        rel_elem = rel_tree.find(f".//rel:Relationship[@Id='{chart_rId}']", namespaces=rel_ns)
        
        if rel_elem is None:
            return None
        
        chart_target = rel_elem.get('Target')
        
        # Строим путь к chart файлу
        if chart_target.startswith('..'):
            parts = slide_rel_path.split('/')
            base_path = '/'.join(parts[:-2])
            clean_target = chart_target.replace('../', '')
            chart_path = f"{base_path}/{clean_target}"
        else:
            chart_path = f"ppt/{chart_target}"
        
        chart_path = chart_path.replace('//', '/')
        return extract_chart_text_by_path(pptx_zip, chart_path, container)
        
    except Exception as e:
        print(f"   ⚠️ Ошибка извлечения диаграммы: {e}")
        return None

# ============================================================
# ТАБЛИЦЫ
# ============================================================

def extract_table_text(table_element):
    """Извлекает текст из таблицы."""
    try:
        rows = []
        all_text = []
        
        for tr in table_element.xpath('.//a:tr', namespaces=NAMESPACES):
            row_cells = []
            for tc in tr.xpath('.//a:tc', namespaces=NAMESPACES):
                cell_texts = []
                for t in tc.xpath('.//a:t', namespaces=NAMESPACES):
                    if t.text and t.text.strip():
                        cell_texts.append(t.text.strip())
                cell_text = ' '.join(cell_texts)
                row_cells.append(cell_text)
                if cell_text:
                    all_text.append(cell_text)
            if row_cells:
                rows.append(row_cells)
        
        return {
            'rows': rows,
            'text': ' | '.join(all_text),
            'is_table': True
        }
    except Exception as e:
        print(f"   ⚠️ Ошибка таблицы: {e}")
        return None

# ============================================================
# ОСНОВНАЯ ФУНКЦИЯ ИЗВЛЕЧЕНИЯ
# ============================================================

def extract_text_blocks(xml_path, pptx_zip=None, slide_rel_path=None):
    """
    Извлекает все текстовые блоки из XML слайда.
    
    Args:
        xml_path: путь к slide.xml
        pptx_zip: ZipFile объект (для диаграмм)
        slide_rel_path: путь к .rels файлу (для диаграмм)
    
    Returns:
        list: блоки с текстом, координатами и форматированием
    """
    if not os.path.exists(xml_path):
        return []
    
    try:
        tree = etree.parse(xml_path)
    except Exception as e:
        print(f"   ⚠️ Ошибка парсинга {xml_path}: {e}")
        return []
    
    # Очищаем кэш координат для нового слайда
    clear_coordinate_cache()
    
    blocks = []
    block_id = 0
    
    # Находим все контейнеры с текстом
    containers = tree.xpath('//p:sp | //a:tbl | //p:graphicFrame', namespaces=NAMESPACES)
    
    for container in containers:
        try:
            # ==================================================
            # 1. ДИАГРАММЫ
            # ==================================================
            if container.tag.endswith('graphicFrame'):
                graphic_data = container.find('.//a:graphicData', namespaces=NAMESPACES)
                if graphic_data is not None and 'chart' in graphic_data.get('uri', ''):
                    # Пропускаем таблицы внутри graphicFrame
                    if container.find('.//a:tbl', namespaces=NAMESPACES) is not None:
                        continue
                    
                    chart_ref = container.find('.//a:chart', namespaces=NAMESPACES)
                    if chart_ref is not None and pptx_zip:
                        chart_rId = chart_ref.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                        chart_data = extract_chart_text(pptx_zip, chart_rId, container, slide_rel_path)
                        
                        if chart_data and chart_data.get('text'):
                            blocks.append({
                                "id": block_id,
                                "text": chart_data['text'],
                                "x": chart_data['x'],
                                "y": chart_data['y'],
                                "w": chart_data['w'],
                                "h": chart_data['h'],
                                "runs": chart_data['runs'],
                                "runs_count": len(chart_data['runs']),
                                "type": "chart"
                            })
                            block_id += 1
                    continue
            
            # ==================================================
            # 2. ТАБЛИЦЫ
            # ==================================================
            if container.tag.endswith('tbl'):
                table_data = extract_table_text(container)
                if table_data and table_data.get('text'):
                    abs_x, abs_y = get_absolute_coordinates(container)
                    blocks.append({
                        "id": block_id,
                        "text": table_data['text'],
                        "x": abs_x / EMU_PER_INCH,
                        "y": abs_y / EMU_PER_INCH,
                        "w": 0,
                        "h": 0,
                        "runs": [{'text': table_data['text'], 'formatting': {}}],
                        "runs_count": 1,
                        "type": "table",
                        "table_data": table_data['rows']
                    })
                    block_id += 1
                continue
            
            # ==================================================
            # 3. ОБЫЧНЫЕ ТЕКСТОВЫЕ БЛОКИ
            # ==================================================
            text_elements = container.xpath('.//a:t', namespaces=NAMESPACES)
            if not text_elements:
                continue
            
            all_runs = []
            full_text_parts = []
            
            for elem in text_elements:
                if not elem.text or not elem.text.strip():
                    continue
                
                full_text_parts.append(elem.text)
                
                # Получаем форматирование
                parent = elem.getparent()
                rPr = parent.find('.//a:rPr', namespaces=NAMESPACES) if parent is not None else None
                
                formatting = {
                    'bold': rPr.get('b') == '1' if rPr is not None else False,
                    'italic': rPr.get('i') == '1' if rPr is not None else False,
                    'underline': rPr.get('u') == 'sng' if rPr is not None else False,
                    'color': None,
                    'font_size': rPr.get('sz') if rPr is not None else None
                }
                
                if rPr is not None:
                    solidFill = rPr.find('.//a:solidFill', namespaces=NAMESPACES)
                    if solidFill is not None:
                        srgbClr = solidFill.find('.//a:srgbClr', namespaces=NAMESPACES)
                        if srgbClr is not None:
                            formatting['color'] = srgbClr.get('val')
                
                all_runs.append({
                    'text': elem.text,
                    'formatting': formatting,
                    'elem_xpath': tree.getpath(elem)
                })
            
            if not all_runs:
                continue
            
            # Координаты и размер
            abs_x, abs_y = get_absolute_coordinates(container)
            w, h = _get_element_size(container)
            
            # Определяем тип блока (заголовок)
            block_type = "shape"
            is_title = False
            
            nvPr = container.find('.//p:nvPr', namespaces=NAMESPACES)
            if nvPr is not None:
                ph = nvPr.find('.//p:ph', namespaces=NAMESPACES)
                if ph is not None and ph.get('type') == 'title':
                    is_title = True
                    block_type = "title"
            
            blocks.append({
                "id": block_id,
                "text": ' '.join(full_text_parts).strip(),
                "x": abs_x / EMU_PER_INCH,
                "y": abs_y / EMU_PER_INCH,
                "w": w / EMU_PER_INCH,
                "h": h / EMU_PER_INCH,
                "runs": all_runs,
                "runs_count": len(all_runs),
                "type": block_type,
                "is_title": is_title
            })
            block_id += 1
            
        except Exception as e:
            print(f"   ⚠️ Ошибка блока {block_id}: {e}")
            continue
    
    # Сортируем блоки по координатам
    blocks.sort(key=lambda b: (b['y'], b['x']))
    
    # Переназначаем ID после сортировки
    for idx, block in enumerate(blocks):
        block['id'] = idx
    
    return blocks


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

def safe_save_xml(tree, filepath):
    try:
        tree.write(filepath, encoding='utf-8', xml_declaration=True, standalone='yes', pretty_print=False)
        return True
    except Exception as e:
        print(f"Ошибка сохранения XML: {e}")
        return False


def apply_translation_safe(xml_path, order, translations_list, blocks_cache, session):
    print(f"   🔧 apply_translation_safe started")
    print(f"   xml_path exists: {os.path.exists(xml_path)}")
    print(f"   order length: {len(order)}")
    print(f"   translations_list length: {len(translations_list)}")
    print(f"   blocks_cache length: {len(blocks_cache)}")
    
    if not os.path.exists(xml_path):
        print(f"   ❌ XML file not found: {xml_path}")
        return False
    
    try:
        tree = etree.parse(xml_path)
        print(f"   ✅ XML parsed successfully")
    except Exception as e:
        print(f"   ❌ XML parse error: {e}")
        return False
    
    # Собираем переводы по ID блока
    translations_by_id = {}
    for i, block_id in enumerate(order):
        if i < len(translations_list) and translations_list[i]:
            t = translations_list[i]
            if isinstance(t, dict):
                t = t.get('translation') or t.get('variant_1_translation') or ''
            if t and isinstance(t, str) and t.strip():
                translations_by_id[block_id] = t
    
    print(f"   📝 translations_by_id: {len(translations_by_id)} записей")
    
    if not translations_by_id:
        print(f"   ⚠️ Нет переводов для применения")
        return False
    
    replaced = 0
    if session not in updated_charts:
        updated_charts[session] = {}
    
    for block in blocks_cache:
        block_id = block['id']
        if block_id not in translations_by_id:
            continue
        
        new_text = translations_by_id[block_id]
        if not new_text or not new_text.strip():
            continue
        
        print(f"   🔄 Применяю перевод для блока {block_id}: {new_text[:50]}...")
        
        # ==================================================
        # 1. ДИАГРАММЫ
        # ==================================================
        if block.get('type') == 'chart' and block.get('chart_file'):
            runs = block.get('runs', [])
            chart_file = block.get('chart_file')
            
            if chart_file not in updated_charts[session]:
                with zipfile.ZipFile(sessions_info[session]['original_path'], 'r') as zf:
                    updated_charts[session][chart_file] = etree.fromstring(zf.read(chart_file))
            chart_tree = updated_charts[session][chart_file]
            
            # Разбиваем с правилом пробелов
            chart_parts = []
            search_pos = 0
            while search_pos < len(new_text):
                pipe_pos = new_text.find('|', search_pos)
                if pipe_pos == -1:
                    part_text = new_text[search_pos:].strip()
                    chart_parts.append({'text': part_text, 'add_space': False})
                    break
                
                part_before = new_text[search_pos:pipe_pos]
                has_left_space = part_before.endswith(' ')
                has_right_space = (pipe_pos + 1 < len(new_text) and new_text[pipe_pos + 1] == ' ')
                clean_text = part_before.strip()
                need_space = has_left_space or has_right_space
                
                chart_parts.append({'text': clean_text, 'add_space': need_space})
                
                search_pos = pipe_pos + 1
                if has_right_space:
                    search_pos += 1
            
            chart_ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                       'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
            
            for i, run in enumerate(runs):
                xpath = run.get('xpath')
                if xpath and i < len(chart_parts):
                    text_to_set = chart_parts[i]['text']
                    if i < len(chart_parts) - 1 and chart_parts[i]['add_space']:
                        text_to_set += ' '
                    for elem in chart_tree.xpath(xpath, namespaces=chart_ns):
                        elem.text = text_to_set if text_to_set else ""
            replaced += 1
            print(f"      ✅ Диаграмма обновлена")
            continue
        
        # ==================================================
        # 2. ТАБЛИЦЫ
        # ==================================================
        if block.get('type') == 'table':
            xpath = block.get('xpath')
            if not xpath and block.get('runs') and len(block['runs']) > 0:
                xpath = block['runs'][0].get('elem_xpath')
            
            if xpath:
                elements = tree.xpath(xpath, namespaces=NAMESPACES)
                if elements:
                    # Для таблиц просто убираем лишние пробелы
                    elements[0].text = re.sub(r'\s+', ' ', new_text.replace('|', ' ')).strip()
                    replaced += 1
                    print(f"      ✅ Таблица обновлена")
            continue
        
        # ==================================================
        # 3. ОБЫЧНЫЕ БЛОКИ
        # ==================================================
        # Получаем xpath из runs если нет прямого xpath
        xpath = block.get('xpath')
        if not xpath and block.get('runs') and len(block['runs']) > 0:
            xpath = block['runs'][0].get('elem_xpath')
        
        if not xpath:
            print(f"      ⚠️ Нет xpath для блока {block_id}")
            continue
        
        # Находим элемент по xpath
        elements = tree.xpath(xpath, namespaces=NAMESPACES)
        if not elements:
            print(f"      ❌ Элемент не найден по xpath: {xpath}")
            continue
        
        elem = elements[0]
        runs = block.get('runs', [])
        
        # Если нет runs - просто заменяем текст
        if not runs:
            elem.text = new_text
            replaced += 1
            print(f"      ✅ Простой текст обновлен")
            continue
        
        # Сложный случай с runs
        runs_count = len(runs)
        reorder_match = re.match(r'^\[([\d,]+)\]\s*(.*)', new_text)
        order_indices = None
        text_after = new_text
        
        if reorder_match:
            order_indices = [int(x) - 1 for x in reorder_match.group(1).split(',')]
            text_after = reorder_match.group(2)
        
        # ==================================================
        # РАЗБИВКА НА ЧАСТИ С ПРАВИЛОМ ПРОБЕЛОВ
        # ==================================================
        parts = []
        if '|' in text_after:
            search_pos = 0
            while search_pos < len(text_after):
                pipe_pos = text_after.find('|', search_pos)
                if pipe_pos == -1:
                    # Последняя часть
                    part_text = text_after[search_pos:].strip()
                    parts.append({'text': part_text, 'add_space': False})
                    break
                
                # Часть до |
                part_before = text_after[search_pos:pipe_pos]
                
                # Проверяем наличие пробелов вокруг |
                has_left_space = part_before.endswith(' ')
                has_right_space = (pipe_pos + 1 < len(text_after) and text_after[pipe_pos + 1] == ' ')
                
                # Делаем trim части
                clean_text = part_before.strip()
                
                # Решаем, нужен ли пробел после этой части
                # Пробел нужен, если были пробелы слева ИЛИ справа от |
                need_space = has_left_space or has_right_space
                
                parts.append({
                    'text': clean_text,
                    'add_space': need_space
                })
                
                # Перемещаемся после |
                search_pos = pipe_pos + 1
                if has_right_space:
                    search_pos += 1  # пропускаем пробел после |
        else:
            # Нет разделителей - один кусок текста на все runs
            for i in range(runs_count):
                parts.append({'text': text_after, 'add_space': False})
        
        # Выравниваем количество частей
        while len(parts) < runs_count:
            parts.append({'text': '', 'add_space': False})
        parts = parts[:runs_count]
        
        # Находим родительский shape
        parent_shape = elem
        while parent_shape is not None:
            if parent_shape.tag.endswith('sp') or parent_shape.tag.endswith('graphicFrame'):
                break
            parent_shape = parent_shape.getparent()
        
        if parent_shape is None:
            print(f"      ⚠️ Не найден родительский shape")
            continue
        
        # Находим все run элементы в shape
        all_paras = parent_shape.xpath('.//a:p', namespaces=NAMESPACES)
        target_runs = []
        for para in all_paras:
            target_runs.extend(para.xpath('.//a:r', namespaces=NAMESPACES))
        
        if len(target_runs) > runs_count:
            target_runs = target_runs[:runs_count]
        
        if not target_runs:
            print(f"      ⚠️ Не найдены run элементы")
            continue
        
        # ==================================================
        # ПРИМЕНЕНИЕ ПЕРЕВОДА С УЧЕТОМ ПРОБЕЛОВ
        # ==================================================
        if order_indices and len(order_indices) == len(target_runs):
            # С перестановкой
            new_runs = [copy.deepcopy(target_runs[idx]) for idx in order_indices if 0 <= idx < len(target_runs)]
            for i, run_elem in enumerate(new_runs):
                if i < len(parts):
                    t_elem = run_elem.find('.//a:t', namespaces=NAMESPACES)
                    if t_elem is not None:
                        text_to_set = parts[i]['text']
                        if i < len(parts) - 1 and parts[i]['add_space']:
                            text_to_set += ' '
                        t_elem.text = text_to_set if text_to_set else ""
            
            # Переставляем run'ы
            parents_map = {}
            for old_run in target_runs:
                p = old_run.getparent()
                if p not in parents_map:
                    parents_map[p] = []
                parents_map[p].append(old_run)
            
            new_runs_by_parent = {}
            run_idx = 0
            for p, old_runs in parents_map.items():
                new_runs_by_parent[p] = [new_runs[run_idx + i] for i in range(len(old_runs))]
                run_idx += len(old_runs)
            
            for p, old_runs in parents_map.items():
                for old_run in old_runs:
                    p.remove(old_run)
                insert_pos = 0
                for new_run in new_runs_by_parent[p]:
                    p.insert(insert_pos, new_run)
                    insert_pos += 1
            
            # Удаление лишних run'ов
            all_runs_in_shape = parent_shape.xpath('.//a:r', namespaces=NAMESPACES)
            for extra_run in all_runs_in_shape[runs_count:]:
                extra_parent = extra_run.getparent()
                if extra_parent is not None:
                    extra_parent.remove(extra_run)
            
            replaced += 1
            print(f"      ✅ Блок с перестановкой обновлен")
        else:
            # Обычное применение без перестановки
            for i, run_elem in enumerate(target_runs):
                if i < len(parts):
                    t_elem = run_elem.find('.//a:t', namespaces=NAMESPACES)
                    if t_elem is not None:
                        text_to_set = parts[i]['text']
                        if i < len(parts) - 1 and parts[i]['add_space']:
                            text_to_set += ' '
                        t_elem.text = text_to_set if text_to_set else ""
            
            # Удаление лишних run'ов
            all_runs_in_shape = parent_shape.xpath('.//a:r', namespaces=NAMESPACES)
            for extra_run in all_runs_in_shape[runs_count:]:
                extra_parent = extra_run.getparent()
                if extra_parent is not None:
                    extra_parent.remove(extra_run)
            
            replaced += 1
            print(f"      ✅ Обычный блок обновлен")
               
    # Сохраняем изменения в XML
    if replaced > 0:
        try:
            tree.write(xml_path, encoding='utf-8', xml_declaration=True, standalone='yes', pretty_print=False)
            print(f"   ✅ XML сохранен. Применено переводов: {replaced}")
            return True
        except Exception as e:
            print(f"   ❌ Ошибка сохранения XML: {e}")
            return False
    else:
        print(f"   ⚠️ Не применено ни одного перевода")
        return False

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
                chart_tree.write(str(chart_path), encoding='utf-8', xml_declaration=True,
                                standalone='yes', pretty_print=False)
        
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
    """Переводит выбранные слайды."""
    print(f"\n{'='*50}")
    print(f"🚀 ЗАПРОС НА ПЕРЕВОД ВЫБРАННЫХ СЛАЙДОВ")
    print(f"{'='*50}")
    
    try:
        if session not in sessions_info:
            return jsonify({"error": "Session not found"}), 404
        
        data = request.get_json()
        api_key = data.get('api_key', USER_SETTINGS.get('api_key', ''))
        model = data.get('model', 'deepseek-reasoner')
        proxy_url = data.get('proxy_url', '')
        target_lang = data.get('target_lang', 'ru')
        enabled_variant_ids = data.get('enabled_variants', None)
        slides_to_translate = data.get('slides', [])  # ← ПРАВИЛЬНОЕ ИМЯ
        
        print(f"   📊 Выбрано слайдов: {len(slides_to_translate)}")
        
        if not api_key:
            return jsonify({"error": "API key required"}), 400
        
        if not slides_to_translate:
            return jsonify({"error": "No slides selected"}), 400
        
        info = sessions_info[session]
        all_slides = slides_list(info['work_dir'])
        
        # Собираем тексты только с выбранных слайдов
        all_texts = []
        with zipfile.ZipFile(info['original_path'], 'r') as zf:
            for slide_idx in slides_to_translate:  # ← ПРАВИЛЬНОЕ ИМЯ
                if slide_idx >= len(all_slides):
                    continue
                blocks = extract_text_blocks(all_slides[slide_idx], zf)
                
                for block in blocks:
                    text = block.get('text', '')
                    runs = block.get('runs', [])
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
        
        if not all_texts:
            return jsonify({"success": False, "error": "No text blocks"}), 400
        
        # Формируем промпт
        print(f"   📏 Формирование промпта...")
        prompt, active_variants = _build_all_slides_prompt(all_texts, target_lang, enabled_variant_ids)
        print(f"   📏 Размер промпта: {len(prompt)} символов")
        print(f"   🎨 Активных вариантов: {len(active_variants)}")
        
        # Вызов DeepSeek
        print(f"   🤖 Вызов DeepSeek API...")
        response = _call_deepseek_single(api_key, model, prompt, proxy_url, session)
        print(f"   📥 Ответ получен, длина: {len(response)} символов")
        
        # Парсим ответ
        print(f"   🔍 Парсинг ответа...")
        translations_matrix = _parse_translations_matrix(response, len(all_texts), len(active_variants))
        print(f"   ✅ Распарсено блоков: {len(translations_matrix)}")
        
        # ========== СОХРАНЯЕМ ПОЛНЫЙ ОТВЕТ В ФАЙЛ ==========
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        response_file = OUT_DIR / f"api_response_{session[:8]}_{timestamp}.json"
        
        full_response = {
            "session_id": session,
            "timestamp": timestamp,
            "slides": slides_to_translate,
            "blocks": all_texts,
            "active_variants": [{"id": v['id'], "name": v['name'], "code": v.get('code', '')} for v in active_variants],
            "translations_matrix": translations_matrix
        }
        
        with open(response_file, 'w', encoding='utf-8') as f:
            json.dump(full_response, f, ensure_ascii=False, indent=2)
        
        print(f"   💾 Полный ответ сохранен в: {response_file}")
        # ==================================================
        
        # Сохраняем переводы в сессию
        if 'translations' not in info:
            info['translations'] = {}
        
        for i, t in enumerate(all_texts):
            slide_idx = t['slide_num'] - 1
            if slide_idx not in info['translations']:
                info['translations'][slide_idx] = {}
            
            block_translations = {}
            if i < len(translations_matrix):
                for v_idx, variant in enumerate(active_variants):
                    if v_idx < len(translations_matrix[i]):
                        block_translations[f"{variant['id']}_translation"] = translations_matrix[i][v_idx]
                        block_translations[f"{variant['id']}_source"] = 'api'
            
            info['translations'][slide_idx][str(t['block_id'])] = block_translations
        
        print(f"   💾 Переводы сохранены в сессию")
        print(f"   ✅ Перевод завершен успешно")
        print(f"{'='*50}\n")
        
        translations_for_json = {}
        for slide_idx, slide_trans in info['translations'].items():
            translations_for_json[str(slide_idx)] = slide_trans

        return jsonify({
            "success": True,
            "total_slides": len(slides),  # ← ИСПРАВЛЕНО
            "translated_blocks": len(translations_matrix),
            "active_variants": active_variants,
            "translations": translations_for_json
        })
        
    except Exception as e:
        print(f"   ❌ ОШИБКА: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500
    
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
        
        session = str(uuid.uuid4())
        original_name = os.path.splitext(file.filename)[0]
        
        # PDF файл (если загружен PDF)
        if file.filename.lower().endswith('.pdf'):
            pdf_path = UPLOAD_DIR / f"{session}.pdf"
            file.save(pdf_path)
            
            sessions_info[session] = {
                'original_name': original_name,
                'original_path': str(pdf_path),
                'total_slides': 1,
                'is_pdf': True
            }
            
            return jsonify({
                "session": session,
                "total_slides": 1,
                "original_name": original_name,
                "is_pdf": True
            })
        
        # PPTX файл
        pptx_path = UPLOAD_DIR / f"{session}.pptx"
        file.save(pptx_path)
        
        # Распаковываем PPTX
        work_dir = WORK_DIR / session
        unzip(pptx_path, work_dir)
        
        # Считаем количество слайдов
        slides = slides_list(work_dir)
        total_slides = len(slides)
        
        # Сохраняем информацию о сессии
        sessions_info[session] = {
            'original_name': original_name,
            'original_path': str(pptx_path),
            'work_dir': str(work_dir),
            'total_slides': total_slides,
            'is_pdf': False,
            'preview_ready': False,
            'preview_pages': 0,
            'created_at': datetime.now()  # ← ДОБАВИТЬ ЭТУ СТРОКУ
        }
        
        # Генерируем превью в фоновом потоке
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
                        import fitz
                        # Исправление для новых версий PyMuPDF
                        if hasattr(fitz, 'Document'):
                            doc = fitz.Document(str(pdf_file))
                        else:
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
    """Возвращает изображение слайда (PNG) для предпросмотра."""
    try:
        info = sessions_info.get(session)
        
        if not info:
            return jsonify({"error": "Session not found"}), 404
        
        # Для PDF
        if info.get('is_pdf'):
            img_path = IMG_DIR / session / f"slide_{slide_idx + 1}.png"
            if img_path.exists():
                return send_file(img_path, mimetype='image/png')
            return jsonify({"error": "Image not found"}), 404
        
        # Для PPTX: ищем в work_dir/preview
        work_dir = Path(info['work_dir'])
        preview_dir = work_dir / "preview"
        img_path = preview_dir / f"slide_{slide_idx + 1}.png"
        
        if img_path.exists():
            return send_file(img_path, mimetype='image/png')
        
        # Запасной вариант: старая папка img
        img_path = IMG_DIR / session / f"slide_{slide_idx + 1}.png"
        if img_path.exists():
            return send_file(img_path, mimetype='image/png')
        
        # Если изображения нет — возвращаем 404, но не генерируем на лету
        return jsonify({"error": "Image not found", "generating": info.get('preview_ready', False)}), 404
        
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
                        import fitz
                        doc = fitz.open(str(pdf_file))
                        for idx, page in enumerate(doc):
                            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                            pix.save(str(preview_dir / f"slide_{idx + 1}.png"))
                        doc.close()
                        print(f"   ✅ Сгенерировано {len(doc)} превью для сессии {session}")
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
    
    success = apply_translation_safe(xml_path, order, translations_list, blocks_for_slide, session)
    
    print(f"Apply result: {success}")
    return jsonify({"success": success})

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


# ============================================================
# ВЫЗОВ DEEPSEEK API
# ============================================================

def _call_deepseek_single(api_key, model, prompt, proxy_url, session_id=None):
    import time
    import ssl
    from requests.adapters import HTTPAdapter
    from urllib3.poolmanager import PoolManager

    class SSLAdapter(HTTPAdapter):
        def init_poolmanager(self, *args, **kwargs):
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            ctx.set_ciphers('DEFAULT@SECLEVEL=1')
            kwargs['ssl_context'] = ctx
            return super().init_poolmanager(*args, **kwargs)

    print(f"   ⏳ Отправка запроса к DeepSeek API...")
    print(f"   📏 Размер промпта: {len(prompt)} символов (~{len(prompt)//4} токенов)")

    proxies = {"https": proxy_url} if proxy_url else None
    start_time = time.time()

    try:
        session = requests.Session()
        session.mount('https://', SSLAdapter())

        response = session.post(
            f"{DEEPSEEK_BASE_URL}/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={
                "model": model,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.3,
                "stream": False
            },
            timeout=300,
            proxies=proxies
        )

        elapsed = time.time() - start_time
        print(f"   ⏱️ Запрос выполнен за {elapsed:.1f} секунд")
        print(f"   📡 Статус ответа: {response.status_code}")

        if response.status_code != 200:
            error_text = response.text[:500]
            print(f"   ❌ API Error {response.status_code}: {error_text}")
            raise Exception(f"API Error {response.status_code}")

        data = response.json()

        usage = data.get('usage', {})
        if usage:
            print(f"   📊 Токенов использовано:")
            print(f"      - промпт: {usage.get('prompt_tokens', 0)}")
            print(f"      - ответ: {usage.get('completion_tokens', 0)}")
            print(f"      - всего: {usage.get('total_tokens', 0)}")

        content = data["choices"][0]["message"]["content"]
        print(f"   📥 Контент получен, длина: {len(content)} символов")

        # Сохраняем ответ в файл
        if session_id:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            response_file = OUT_DIR / f"api_response_{session_id[:8]}_{timestamp}.txt"
            with open(response_file, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"   💾 Ответ сохранен в: {response_file}")

        return content

    except Exception as e:
        print(f"   ❌ Ошибка API: {e}")
        raise


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


def _parse_translations_matrix(response_text, expected_blocks, expected_variants):
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

    # Дополнение до ожидаемого размера
    while len(matrix) < expected_blocks:
        matrix.append([''] * expected_variants)

    for i in range(len(matrix)):
        if not isinstance(matrix[i], list):
            matrix[i] = [str(matrix[i])] if matrix[i] else []
        while len(matrix[i]) < expected_variants:
            matrix[i].append('')

    return matrix[:expected_blocks]


# ============================================================
# СОХРАНЕНИЕ MAPPING И ЗАГРУЗКА ИЗ ФАЙЛА
# ============================================================

@app.route('/api/translate_with_filters/<session>', methods=['POST'])
def translate_with_filters(session):
    """Перевод выбранных слайдов с фильтрацией блоков."""
    try:
        if session not in sessions_info:
            return jsonify({"error": "Session not found"}), 404

        data = request.get_json()
        api_key = data.get('api_key', USER_SETTINGS.get('api_key', ''))
        model = data.get('model', 'deepseek-reasoner')
        proxy_url = data.get('proxy_url', '')
        slides_to_translate = data.get('slides', [])
        enabled_variant_ids = data.get('enabled_variants', [])

        if not api_key:
            return jsonify({"error": "API key required"}), 400

        all_texts, mapping = collect_filtered_texts(session, slides_to_translate, skip_russian=True)

        if not all_texts:
            return jsonify({"error": "No blocks to translate"}), 400

        # Сохраняем mapping для будущей загрузки
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        mapping_file = OUT_DIR / f"mapping_{session[:8]}_{timestamp}.json"
        active_variants = get_variants_list()
        with open(mapping_file, 'w', encoding='utf-8') as f:
            json.dump({
                'mapping': mapping,
                'active_variants': active_variants
            }, f, ensure_ascii=False, indent=2)

        prompt, _ = _build_all_slides_prompt(all_texts, 'ru', enabled_variant_ids)

        response = _call_deepseek_single(api_key, model, prompt, proxy_url, session)

        translations_matrix = _parse_translations_matrix(response, len(all_texts), len(active_variants))

        # Сохраняем переводы
        info = sessions_info[session]
        if 'translations' not in info:
            info['translations'] = {}

        for i, m in enumerate(mapping):
            if i >= len(translations_matrix):
                break
            slide_idx = m['slide_num'] - 1
            block_id = str(m['block_id'])
            if slide_idx not in info['translations']:
                info['translations'][slide_idx] = {}
            block_trans = {}
            row = translations_matrix[i]
            for v_idx, variant in enumerate(active_variants):
                if v_idx < len(row) and row[v_idx]:
                    block_trans[f"{variant['id']}_translation"] = row[v_idx]
                    block_trans[f"{variant['id']}_source"] = 'api'
            if block_trans:
                info['translations'][slide_idx][block_id] = block_trans

        # Приводим все ключи слайдов к строкам
        translations_serializable = {}
        for slide_idx, slide_trans in info['translations'].items():
            translations_serializable[str(slide_idx)] = slide_trans

        return jsonify({
            "success": True,
            "translated_blocks": len(translations_matrix),
            "mapping_file": str(mapping_file),
            "translations": translations_serializable
        })

    except Exception as e:
        print(f"❌ Ошибка перевода: {e}")
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route('/api/load_response_from_file', methods=['POST'])
def load_response_from_file():
    """Загружает ответ API из файла и применяет к сессии."""
    try:
        data = request.get_json()
        session = data.get('session')
        file_path = data.get('file_path')
        mapping_path = data.get('mapping_path')

        if session not in sessions_info:
            return jsonify({"error": "Session not found"}), 404

        with open(file_path, 'r', encoding='utf-8') as f:
            response_text = f.read()

        with open(mapping_path, 'r', encoding='utf-8') as f:
            mapping_data = json.load(f)

        mapping = mapping_data['mapping']
        active_variants = mapping_data['active_variants']

        expected_blocks = len(mapping)
        expected_variants = len(active_variants)

        translations_matrix = _parse_translations_matrix(response_text, expected_blocks, expected_variants)

        info = sessions_info[session]
        if 'translations' not in info:
            info['translations'] = {}

        saved = 0
        for i, m in enumerate(mapping):
            if i >= len(translations_matrix):
                break
            slide_idx = m['slide_num'] - 1
            block_id = str(m['block_id'])
            if slide_idx not in info['translations']:
                info['translations'][slide_idx] = {}
            block_trans = {}
            row = translations_matrix[i]
            for v_idx, variant in enumerate(active_variants):
                if v_idx < len(row) and row[v_idx]:
                    block_trans[f"{variant['id']}_translation"] = row[v_idx]
                    block_trans[f"{variant['id']}_source"] = 'file'
            if block_trans:
                info['translations'][slide_idx][block_id] = block_trans
                saved += 1

        translations_serializable = {}
        for slide_idx, slide_trans in info['translations'].items():
            translations_serializable[str(slide_idx)] = slide_trans

        return jsonify({
            "success": True,
            "saved_blocks": saved,
            "translations": translations_serializable
        })

    except Exception as e:
        print(f"❌ Ошибка загрузки из файла: {e}")
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route('/api/translate/all_single/<session>', methods=['POST'])
def translate_all_single(session):
    print(f"\n{'='*50}")
    print(f"🚀 ЗАПРОС НА ПЕРЕВОД ВСЕЙ ПРЕЗЕНТАЦИИ")
    print(f"   Session: {session}")
    print(f"{'='*50}\n")
    
    try:
        if session not in sessions_info:
            return jsonify({"error": "Session not found"}), 404
        
        data = request.get_json()
        api_key = data.get('api_key', USER_SETTINGS.get('api_key', ''))
        enabled_variant_ids = data.get('enabled_variants', None)
        
        print(f"   🔑 API key: {'***' if api_key else 'НЕТ'}")
        print(f"   📊 Включенные варианты: {enabled_variant_ids}")
        
        if not api_key:
            return jsonify({"error": "API key required"}), 400
        
        print("   📦 Сбор текстов...")
        all_texts = collect_all_texts_with_meta(session)
        print(f"   📊 Собрано блоков: {len(all_texts)}")
        
        if not all_texts:
            return jsonify({"success": False, "error": "No text blocks"}), 400
        
        print("   📝 Формирование промпта...")
        prompt, active_variants = _build_all_slides_prompt(
            all_texts, 
            data.get('target_lang', 'ru'),
            enabled_variant_ids
        )
        print(f"   📏 Размер промпта: {len(prompt)} символов")
        print(f"   🎨 Активных вариантов: {len(active_variants)}")
        
        print("   🤖 Вызов DeepSeek API...")
        response = _call_deepseek_single(
            api_key, 
            data.get('model', 'deepseek-reasoner'),
            prompt, 
            data.get('proxy_url', '')
        )
        print(f"   📥 Ответ получен, длина: {len(response)} символов")
        
        print("   🔍 Парсинг ответа...")
        translations_matrix = _parse_translations_matrix(response, len(all_texts), len(active_variants))
        print(f"   ✅ Распарсено блоков: {len(translations_matrix)}")
        
        # Сохраняем переводы
        info = sessions_info[session]
        if 'translations' not in info:
            info['translations'] = {}
        
        for i, t in enumerate(all_texts):
            slide_idx = t['slide_num'] - 1
            if slide_idx not in info['translations']:
                info['translations'][slide_idx] = {}
            
            block_translations = {}
            if i < len(translations_matrix):
                for v_idx, variant in enumerate(active_variants):
                    if v_idx < len(translations_matrix[i]):
                        block_translations[f"{variant['id']}_translation"] = translations_matrix[i][v_idx]
                        block_translations[f"{variant['id']}_source"] = 'api'
            
            info['translations'][slide_idx][str(t['block_id'])] = block_translations
        
        print(f"   💾 Переводы сохранены в сессию")
        print(f"   📊 Слайдов с переводами: {len(info['translations'])}")
        
        translations_for_json = {}
        for slide_idx, slide_trans in info['translations'].items():
            translations_for_json[str(slide_idx)] = slide_trans

        return jsonify({
            "success": True,
            "total_slides": len(set(t['slide_num'] for t in all_texts)),  # ← ИСПРАВЛЕНО
            "translated_blocks": len(translations_matrix),
            "active_variants": active_variants,
            "translations": translations_for_json
        })
        
    except Exception as e:
        print(f"   ❌ ОШИБКА: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

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
            verify=False
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
    try:
        if session not in sessions_info:
            return jsonify({"error": "Session not found"}), 404
        
        data = request.get_json()
        api_key = data.get('api_key', USER_SETTINGS.get('api_key', ''))
        model = data.get('model', USER_SETTINGS.get('model', 'deepseek-reasoner'))
        proxy_url = data.get('proxy_url', USER_SETTINGS.get('proxy_url', ''))
        target_lang = data.get('target_lang', 'ru')
        enabled_variant_ids = data.get('enabled_variants', None)
        
        if not api_key:
            return jsonify({"error": "API key required"}), 400
        
        info = sessions_info[session]
        slides = slides_list(info['work_dir'])
        if slide_idx >= len(slides):
            return jsonify({"error": "Slide not found"}), 404
        
        with zipfile.ZipFile(info['original_path'], 'r') as zf:
            blocks = extract_text_blocks(slides[slide_idx], zf)
        
        texts = []
        texts_meta = []
        for i, b in enumerate(blocks, 1):
            t = b.get('text', '')
            if b.get('runs') and len(b['runs']) > 1:
                t = ' | '.join([r.get('text', '') for r in b['runs']])
            texts.append(t)
            texts_meta.append({
                'slide_num': 1,
                'block_num': i,
                'block_id': b['id'],
                'text': t,
                'has_runs': len(b.get('runs', [])) > 1,
                'runs_count': len(b.get('runs', []))
            })
        
        # ИСПРАВЛЕНО: принимаем кортеж
        prompt, active_variants = _build_all_slides_prompt(
            texts_meta,
            target_lang,
            enabled_variant_ids
        )
        
        response = _call_deepseek_single(api_key, model, prompt, proxy_url, session)
        translations_matrix = _parse_translations_matrix(response, len(texts), len(active_variants))
        
        # Сохраняем переводы
        if 'translations' not in info:
            info['translations'] = {}

        saved_slides = set()
        for i, t in enumerate(texts_meta):
            slide_idx = t['slide_num'] - 1
            saved_slides.add(slide_idx)
            
            if slide_idx not in info['translations']:
                info['translations'][slide_idx] = {}
            
            block_translations = {}
            if i < len(translations_matrix):
                for v_idx, variant in enumerate(active_variants):
                    if v_idx < len(translations_matrix[i]):
                        block_translations[f"{variant['id']}_translation"] = translations_matrix[i][v_idx]
                        block_translations[f"{variant['id']}_source"] = 'api'
            
            info['translations'][slide_idx][str(t['block_id'])] = block_translations

        print(f"   💾 Переводы сохранены для слайдов: {sorted(saved_slides)}")
        print(f"   📊 Всего слайдов с переводами: {len(saved_slides)}")
        
        translations_for_json = {}
        for slide_idx, slide_trans in info['translations'].items():
            translations_for_json[str(slide_idx)] = slide_trans

        return jsonify({
            "success": True,
            "total_slides": len(slides_to_translate),
            "translated_blocks": len(translations_matrix),
            "active_variants": active_variants,
            "translations": translations_for_json
        })
        
        
    except Exception as e:
        print(f"❌ Ошибка перевода слайда: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500
    
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
    
@app.route('/api/list_response_files', methods=['GET'])
def list_response_files():
    """Возвращает список сохраненных ответов API и mapping файлов."""
    try:
        responses = []
        mappings = []
        
        for f in OUT_DIR.glob("api_response_*.txt"):
            responses.append({
                "name": f.name,
                "path": str(f),
                "size": f.stat().st_size,
                "modified": datetime.fromtimestamp(f.stat().st_mtime).isoformat()
            })
        
        for f in OUT_DIR.glob("mapping_*.json"):
            mappings.append({
                "name": f.name,
                "path": str(f),
                "size": f.stat().st_size,
                "modified": datetime.fromtimestamp(f.stat().st_mtime).isoformat()
            })
        
        responses.sort(key=lambda x: x['modified'], reverse=True)
        mappings.sort(key=lambda x: x['modified'], reverse=True)
        
        return jsonify({
            "success": True,
            "responses": responses[:50],
            "mappings": mappings[:50]
        })
    except Exception as e:
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
    app.run(host=APP_HOST, port=APP_PORT, debug=True)