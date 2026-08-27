"""
Главный модуль приложения.
"""

from cli import show_welcome, get_directory, get_user_choice
from comparator import compare_directories
from image_converter import convert_images_in_folder, get_conversion_stats
from word_converter import convert_word_documents_in_folder, needs_word_conversion
import config


def main():
    """Основная функция - запрос папок и сравнение."""
    show_welcome()

    dir1 = get_directory("Введите путь к ПАПКЕ 1: ")
    dir2 = get_directory("Введите путь к ПАПКЕ 2: ")

    # 1. КОНВЕРТАЦИЯ ИЗОБРАЖЕНИЙ
    if config.CONVERT_IMAGES_TO_PNG:
        print("\n--- КОНВЕРТАЦИЯ ИЗОБРАЖЕНИЙ В PNG (С ЗАМЕНОЙ) ---")
        print("Поиск изображений для конвертации...")

        converted1, skipped1, errors1 = convert_images_in_folder(dir1)
        converted2, skipped2, errors2 = convert_images_in_folder(dir2)

        total1, saved1, orig1, new1 = get_conversion_stats(converted1)
        total2, saved2, orig2, new2 = get_conversion_stats(converted2)

        total_files = total1 + total2
        total_saved = saved1 + saved2
        total_skipped = skipped1 + skipped2

        if total_files > 0:
            print(f"\n✅ Конвертация изображений завершена!")
            print(f"   📸 Всего сконвертировано: {total_files} файлов")
            print(f"   💾 Сэкономлено места: ~{total_saved / 1024:.1f} КБ")
            if config.VERBOSE_OUTPUT:
                for rel_path, info in converted1.items():
                    print(f"   📸 {rel_path} → PNG (экономия {info['saved_percent']}%)")
                for rel_path, info in converted2.items():
                    if rel_path not in converted1:
                        print(f"   📸 {rel_path} → PNG (экономия {info['saved_percent']}%)")
        else:
            print("ℹ️ Изображения для конвертации не найдены или все уже в PNG/TIFF.")

        if total_skipped > 0:
            print(f"ℹ️ Пропущено файлов (уже PNG/TIFF): {total_skipped}")
        if errors1 + errors2 > 0:
            print(f"⚠️ Ошибок при конвертации: {errors1 + errors2}")

    # 2. КОНВЕРТАЦИЯ WORD → ODT
    if config.CONVERT_WORD_TO_ODT:
        print("\n--- КОНВЕРТАЦИЯ WORD → ODT ---")

        # Проверяем, есть ли файлы для конвертации
        import os
        from pathlib import Path

        doc_files = list(Path(dir1).glob("*.doc")) + list(Path(dir1).glob("*.docx"))
        doc_files += list(Path(dir2).glob("*.doc")) + list(Path(dir2).glob("*.docx"))

        if doc_files:
            print(f"📁 Найдено {len(doc_files)} документов Word")

            # Конвертируем в первой папке
            stats1 = convert_word_documents_in_folder(dir1)

            # Конвертируем во второй папке
            stats2 = convert_word_documents_in_folder(dir2)

            total_success = stats1.get("success", 0) + stats2.get("success", 0)
            total_errors = stats1.get("errors", 0) + stats2.get("errors", 0)

            if total_success > 0:
                print(f"\n✅ Конвертация Word завершена!")
                print(f"   📄 Успешно сконвертировано: {total_success} файлов")
            if total_errors > 0:
                print(f"   ❌ Ошибок: {total_errors}")
        else:
            print("ℹ️ Документы Word для конвертации не найдены.")

    # 3. ВЫПОЛНЯЕМ СРАВНЕНИЕ
    compare_directories(dir1, dir2)


def startup():
    """Управляет циклом программы: первый запуск сразу, затем по запросу."""
    first_run = True

    while True:
        if first_run:
            main()
            first_run = False
        else:
            print("\n" + "-"*40)
            if get_user_choice():
                main()
            else:
                print("Выход из программы.")
                break


if __name__ == "__main__":
    startup()