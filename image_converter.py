"""
Модуль для конвертации изображений в PNG с заменой исходных файлов.
Использует библиотеку PIL (Pillow).
"""

import os
from PIL import Image
import config


def convert_image_to_png(source_path, compression_level=None):
    """
    Конвертирует изображение в PNG и ЗАМЕНЯЕТ исходный файл.
    Возвращает True в случае успеха, False при ошибке.
    """
    if compression_level is None:
        compression_level = config.PNG_COMPRESSION_LEVEL

    # Создаём временный путь для нового файла
    temp_path = source_path + '.temp.png'

    try:
        with Image.open(source_path) as img:
            # Конвертируем в RGBA, если изображение имеет прозрачность
            if img.mode in ('RGBA', 'LA', 'P'):
                img = img.convert('RGBA')
            else:
                img = img.convert('RGB')

            # Сохраняем во временный файл
            img.save(temp_path, 'PNG', compress_level=compression_level)

            # Получаем размеры до/после
            original_size = os.path.getsize(source_path)
            new_size = os.path.getsize(temp_path)

            # Удаляем исходный файл
            os.remove(source_path)

            # Переименовываем временный файл в исходное имя с расширением .png
            new_path = os.path.splitext(source_path)[0] + '.png'
            os.rename(temp_path, new_path)

            return True, original_size, new_size, new_path

    except Exception as e:
        # В случае ошибки удаляем временный файл, если он создался
        if os.path.exists(temp_path):
            os.remove(temp_path)
        print(f"  ⚠️ Ошибка конвертации {source_path}: {e}")
        return False, 0, 0, None


def convert_images_in_folder(base_dir):
    """
    Находит все изображения в папке (кроме TIFF и PNG)
    и конвертирует их в PNG, заменяя исходные файлы.

    Возвращает словарь с информацией о сконвертированных файлах.
    """
    converted_files = {}
    skipped_count = 0
    error_count = 0

    # Обходим все папки рекурсивно
    for root, dirs, filenames in os.walk(base_dir):
        for filename in filenames:
            ext = os.path.splitext(filename)[1].lower()

            # 1. Проверяем, нужно ли пропустить (PNG, TIFF)
            if ext in config.SKIP_FORMATS:
                skipped_count += 1
                continue

            # 2. Проверяем, является ли файл изображением (JPG, BMP, GIF и т.д.)
            if ext not in {'.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.ico'}:
                # Не изображение - пропускаем
                continue

            # Полный путь к файлу
            src_path = os.path.join(root, filename)

            # 3. Дополнительная проверка на случай, если файл уже PNG
            if os.path.splitext(src_path)[1].lower() == '.png':
                skipped_count += 1
                continue

            # Конвертируем с заменой
            success, old_size, new_size, new_path = convert_image_to_png(src_path)

            if success:
                rel_path = os.path.relpath(src_path, base_dir)
                converted_files[rel_path] = {
                    'original': src_path,
                    'converted': new_path,
                    'original_size': old_size,
                    'new_size': new_size,
                    'saved_bytes': old_size - new_size,
                    'saved_percent': round((1 - new_size/old_size) * 100, 1) if old_size > 0 else 0
                }
            else:
                error_count += 1

    return converted_files, skipped_count, error_count


def needs_conversion(filename):
    """Проверяет, нужно ли конвертировать файл."""
    ext = os.path.splitext(filename)[1].lower()
    return ext not in config.SKIP_FORMATS and ext in {'.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.ico'}


def get_conversion_stats(converted_files):
    """Возвращает статистику по конвертации."""
    if not converted_files:
        return 0, 0, 0, 0

    total_files = len(converted_files)
    total_saved_bytes = sum(info['saved_bytes'] for info in converted_files.values() if info['saved_bytes'] > 0)
    total_original = sum(info['original_size'] for info in converted_files.values())
    total_new = sum(info['new_size'] for info in converted_files.values())

    return total_files, total_saved_bytes, total_original, total_new