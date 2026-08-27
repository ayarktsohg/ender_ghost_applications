"""
Модуль для работы с пользовательским вводом/выводом.
"""

import os
import sys


def get_directory(prompt):
    """Запрашивает у пользователя путь к папке и проверяет его существование."""
    while True:
        path = input(prompt).strip()
        if os.path.isdir(path):
            return path
        print("Ошибка: Папка не найдена. Попробуйте снова.")


def show_welcome():
    """Показывает приветственное сообщение."""
    print("\n" + "="*60)
    print("   СРАВНИВАТЕЛЬ ПАПОК v2.0")
    print("="*60)
    print("Программа рекурсивно сравнивает две папки,")
    print("включая все вложенные папки и файлы.")
    print("="*60 + "\n")


def get_user_choice(prompt="Продолжить поиск? (1 - да, 0 - выход): "):
    """Запрашивает у пользователя выбор (1/0)."""
    while True:
        try:
            choice = input(prompt).strip()
            if choice == "1":
                return True
            elif choice == "0":
                return False
            else:
                print("Ошибка: введите 1 или 0.")
        except KeyboardInterrupt:
            print("\nВыход из программы.")
            sys.exit(0)