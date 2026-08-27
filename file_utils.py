"""
Модуль для работы с файловой системой: обход папок, получение информации.
"""

import os
from pathlib import Path


def get_all_files_info(directory):
    """
    Рекурсивно обходит все папки и создает словарь:
    {относительный_путь_к_файлу: (полный_путь, размер_в_байтах)}
    """
    files = {}
    for root, dirs, filenames in os.walk(directory):
        for filename in filenames:
            full_path = os.path.join(root, filename)
            rel_path = os.path.relpath(full_path, directory)
            size = os.path.getsize(full_path)
            files[rel_path] = (full_path, size)
    return files


def get_all_dirs_info(directory):
    """
    Рекурсивно собирает информацию о всех папках:
    {относительный_путь_к_папке: список_файлов_в_ней}
    """
    dirs = {}
    for root, dirnames, filenames in os.walk(directory):
        rel_path = os.path.relpath(root, directory)
        if rel_path == '.':
            rel_path = ''  # Корневая папка
        dirs[rel_path] = sorted(filenames)
    return dirs


def is_image_file(filename):
    """Проверяет, является ли файл изображением по расширению."""
    image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', '.webp', '.ico'}
    ext = os.path.splitext(filename)[1].lower()
    return ext in image_extensions


def get_file_extension(filename):
    """Возвращает расширение файла в нижнем регистре."""
    return os.path.splitext(filename)[1].lower()


def create_directory(path):
    """Создаёт папку, если её нет."""
    if not os.path.exists(path):
        os.makedirs(path)
        return True
    return False