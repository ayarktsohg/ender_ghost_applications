"""
Модуль сравнения папок и файлов.
"""

import os
import hashlib
from file_utils import get_all_files_info, get_all_dirs_info
import config


def get_file_hash(filepath, chunk_size=8192):
    """Вычисляет MD5-хеш файла."""
    hasher = hashlib.md5()
    with open(filepath, 'rb') as f:
        for chunk in iter(lambda: f.read(chunk_size), b''):
            hasher.update(chunk)
    return hasher.hexdigest()


def compare_directories(dir1, dir2, convert_images=False):
    """
    Сравнивает две папки рекурсивно:
    - структуру папок (названия)
    - содержимое папок (списки файлов)
    - каждый файл по размеру (или по хешу, если включено)
    """
    print("\n" + "=" * 60)
    print("СРАВНЕНИЕ ПАПОК (РЕКУРСИВНОЕ)")
    print("=" * 60)

    files1 = get_all_files_info(dir1)
    files2 = get_all_files_info(dir2)
    dirs1 = get_all_dirs_info(dir1)
    dirs2 = get_all_dirs_info(dir2)

    has_differences = False

    # 1. Сравниваем структуру папок
    all_dirs = set(dirs1.keys()) | set(dirs2.keys())
    dir_diffs = False

    print("\n--- СТРУКТУРА ПАПОК ---")
    for dir_name in sorted(all_dirs):
        display_name = dir_name or 'корневая'
        if dir_name not in dirs1:
            print(f"❌ Папка '{display_name}': отсутствует в ПАПКЕ 1")
            has_differences = True
            dir_diffs = True
        elif dir_name not in dirs2:
            print(f"❌ Папка '{display_name}': отсутствует в ПАПКЕ 2")
            has_differences = True
            dir_diffs = True
        else:
            if dirs1[dir_name] != dirs2[dir_name]:
                print(f"⚠️ Папка '{display_name}': содержимое различается")
                set1 = set(dirs1[dir_name])
                set2 = set(dirs2[dir_name])
                only_in_1 = set1 - set2
                only_in_2 = set2 - set1
                if only_in_1:
                    print(f"   Только в ПАПКЕ 1: {', '.join(only_in_1)}")
                if only_in_2:
                    print(f"   Только в ПАПКЕ 2: {', '.join(only_in_2)}")
                has_differences = True
                dir_diffs = True

    if not dir_diffs:
        print("✅ Структура папок совпадает")

    # 2. Сравниваем файлы
    all_files = set(files1.keys()) | set(files2.keys())
    file_diffs = False

    print("\n--- ФАЙЛЫ ---")
    count = 0
    for rel_path in sorted(all_files):
        if count >= config.MAX_DIFFERENCES_DISPLAY:
            print(f"... и ещё {len(all_files) - count} различий (ограничение вывода)")
            break

        if rel_path not in files1:
            print(f"❌ {rel_path}: отсутствует в ПАПКЕ 1")
            has_differences = True
            file_diffs = True
            count += 1
        elif rel_path not in files2:
            print(f"❌ {rel_path}: отсутствует в ПАПКЕ 2")
            has_differences = True
            file_diffs = True
            count += 1
        else:
            _, size1 = files1[rel_path]
            _, size2 = files2[rel_path]

            if config.COMPARE_BY_HASH:
                # Сравниваем по хешу
                hash1 = get_file_hash(files1[rel_path][0])
                hash2 = get_file_hash(files2[rel_path][0])
                if hash1 != hash2:
                    print(f"⚠️ {rel_path}: содержимое различается (хеши: {hash1[:8]} vs {hash2[:8]})")
                    has_differences = True
                    file_diffs = True
                    count += 1
            else:
                # Сравниваем по размеру
                if size1 != size2:
                    print(f"⚠️ {rel_path}: размер различается ({size1} vs {size2} байт)")
                    has_differences = True
                    file_diffs = True
                    count += 1

    if not file_diffs:
        print("✅ Все файлы совпадают")

    # Итог
    print("\n" + "=" * 60)
    if not has_differences:
        print("✅ ВСЕ ПАПКИ И ФАЙЛЫ ПОЛНОСТЬЮ СОВПАДАЮТ!")
    else:
        print("❌ ОБНАРУЖЕНЫ РАЗЛИЧИЯ (смотрите выше)")
    print("=" * 60)

    return has_differences