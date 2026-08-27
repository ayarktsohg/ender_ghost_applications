"""
Модуль для конвертации документов Word (DOC, DOCX) в OpenDocument (ODT).
Использует VBA-макросы через COM-интерфейс Microsoft Word.
Требуется установленный Microsoft Word.
"""

import os
import time
import json
import glob
import shutil
import tempfile
import win32com.client as win32
from pathlib import Path
import config


class WordToODTConverter:
    """
    Конвертер Word-документов в ODT с использованием VBA-макросов.
    Обрабатывает файлы .doc и .docx последовательно.
    """

    def __init__(self):
        self.word_app = None
        self.results = []
        self.errors = []
        self.temp_folder = os.path.join(tempfile.gettempdir(), "word_converter_temp")
        self.processed_count = 0
        self.success_count = 0
        self.error_count = 0

    def _connect_word(self):
        """Подключается к Microsoft Word."""
        try:
            self.word_app = win32.Dispatch("Word.Application")
            self.word_app.Visible = False
            self.word_app.DisplayAlerts = 0
            return True
        except Exception as e:
            print(f"  ❌ Ошибка подключения к Word: {e}")
            return False

    def _disconnect_word(self):
        """Закрывает Microsoft Word."""
        try:
            if self.word_app:
                self.word_app.Quit()
                self.word_app = None
        except:
            pass

    def _get_latest_result_file(self, timeout=45):
        """
        Ожидает появления JSON-файла результата от макроса.
        Возвращает путь к файлу или None.
        """
        start_time = time.time()

        # Сначала проверяем переменную среды
        try:
            env_var = os.environ.get("LAST_CONVERT_RESULT")
            if env_var and os.path.exists(env_var) and env_var.endswith('.json'):
                # Проверяем, что файл содержит валидный JSON
                try:
                    with open(env_var, 'r', encoding='utf-8-sig') as f:
                        json.load(f)
                    return env_var
                except:
                    # Если JSON невалидный, удаляем файл
                    try:
                        os.remove(env_var)
                    except:
                        pass
        except:
            pass

        # Ищем файлы в папке
        while time.time() - start_time < timeout:
            pattern = os.path.join(self.temp_folder, "*.json")
            files = glob.glob(pattern)

            if files:
                # Берём самый свежий файл
                latest = max(files, key=os.path.getctime)
                time.sleep(0.3)

                if os.path.getsize(latest) > 0:
                    # Проверяем, что файл содержит валидный JSON
                    try:
                        with open(latest, 'r', encoding='utf-8-sig') as f:
                            json.load(f)
                        return latest
                    except (json.JSONDecodeError, UnicodeDecodeError):
                        # Файл не является валидным JSON - удаляем его
                        try:
                            os.remove(latest)
                        except:
                            pass
                        continue

            time.sleep(0.3)

        return None

    def _ensure_temp_folder(self):
        """Создаёт временную папку, если её нет."""
        if not os.path.exists(self.temp_folder):
            os.makedirs(self.temp_folder)

    def _clean_temp_folder(self):
        """Очищает временную папку от старых файлов."""
        try:
            if os.path.exists(self.temp_folder):
                for f in glob.glob(os.path.join(self.temp_folder, "*.json")):
                    try:
                        os.remove(f)
                    except:
                        pass
        except:
            pass

    def convert_document(self, file_path, macro_name="ConvertToODTAndReport",
                         delete_original=None):
        """
        Конвертирует один документ Word в ODT.

        Аргументы:
            file_path: путь к файлу .doc/.docx
            macro_name: имя макроса в Normal.dotm
            delete_original: удалять ли исходный файл

        Возвращает:
            dict с результатом конвертации
        """
        if delete_original is None:
            delete_original = config.DELETE_ORIGINAL_AFTER_CONVERT

        if not self.word_app:
            return None

        result = {
            "file": file_path,
            "status": "PENDING",
            "original_name": os.path.basename(file_path),
            "odt_name": None,
            "odt_path": None,
            "error": None,
            "deleted": False
        }

        self._ensure_temp_folder()
        # Очищаем старые JSON-файлы перед конвертацией
        self._clean_temp_folder()

        doc = None
        try:
            # Открываем документ
            doc = self.word_app.Documents.Open(file_path)
            time.sleep(1)

            # Запускаем макрос
            self.word_app.Run(macro_name)

            # Ждём появления JSON-файла результата
            json_file = self._get_latest_result_file(timeout=45)

            if json_file:
                try:
                    # Читаем JSON с поддержкой UTF-8 (с BOM и без)
                    with open(json_file, 'r', encoding='utf-8-sig') as f:
                        data = json.load(f)

                    # Обновляем результат из JSON
                    result.update(data)
                    result["original_name"] = os.path.basename(file_path)

                    # Проверяем статус
                    if data.get("status") == "SUCCESS":
                        # Проверяем, создался ли ODT-файл
                        odt_path = data.get("odt_path")
                        odt_file = data.get("odt_file")

                        if odt_path and os.path.exists(odt_path):
                            result["odt_path"] = odt_path
                            result["odt_name"] = odt_file or os.path.basename(odt_path)

                            # Удаляем исходный файл, если включено
                            if delete_original:
                                try:
                                    os.remove(file_path)
                                    result["deleted"] = True
                                except Exception as e:
                                    result["delete_error"] = str(e)
                        else:
                            # ODT не создан, хотя макрос сообщил об успехе
                            result["status"] = "ERROR"
                            result["error"] = "ODT-файл не создан (путь не найден)"

                    elif data.get("status") == "ERROR":
                        # Макрос вернул ошибку
                        result["error"] = data.get("error", "Неизвестная ошибка в макросе")

                    # Удаляем JSON-файл после чтения
                    try:
                        os.remove(json_file)
                    except:
                        pass

                except json.JSONDecodeError as e:
                    result["status"] = "ERROR"
                    result["error"] = f"Ошибка парсинга JSON: {e}"
                    try:
                        os.remove(json_file)
                    except:
                        pass
                except UnicodeDecodeError as e:
                    result["status"] = "ERROR"
                    result["error"] = f"Ошибка кодировки JSON: {e}"
                    try:
                        os.remove(json_file)
                    except:
                        pass
                except Exception as e:
                    result["status"] = "ERROR"
                    result["error"] = f"Ошибка чтения JSON: {e}"
                    try:
                        os.remove(json_file)
                    except:
                        pass
            else:
                # JSON-файл не появился за отведённое время
                result["status"] = "TIMEOUT"
                result["error"] = ("Превышено время ожидания ответа от макроса. "
                                   "Проверьте, что макрос 'ConvertToODTAndReport' "
                                   "существует в Normal.dotm и разрешён для выполнения.")

            # Закрываем документ (без сохранения, т.к. макрос уже сохранил как ODT)
            try:
                if doc:
                    doc.Close(SaveChanges=False)
            except:
                pass

        except Exception as e:
            result["status"] = "ERROR"
            result["error"] = str(e)
            try:
                if doc:
                    doc.Close(SaveChanges=False)
            except:
                pass

        # Сохраняем результат
        self.results.append(result)
        self.processed_count += 1

        if result.get("status") == "SUCCESS":
            self.success_count += 1
        else:
            self.error_count += 1

        return result

    def convert_folder(self, folder_path, extensions=None, macro_name="ConvertToODTAndReport",
                      verbose=None):
        """
        Конвертирует все документы Word в папке.
        """
        if verbose is None:
            verbose = config.VERBOSE_OUTPUT

        if extensions is None:
            extensions = ['.doc', '.docx']

        # Собираем все файлы
        files = []
        for ext in extensions:
            files.extend(Path(folder_path).glob(f"*{ext}"))

        if not files:
            if verbose:
                print(f"ℹ️ В папке {folder_path} не найдено документов Word")
            return self.results

        if verbose:
            print(f"📁 Найдено {len(files)} документов для конвертации")
            print("=" * 60)

        # Подключаемся к Word
        if not self._connect_word():
            print("❌ Не удалось подключиться к Word")
            return self.results

        try:
            for i, file in enumerate(files, 1):
                if verbose:
                    print(f"\n--- [{i}/{len(files)}] ---")
                    print(f"📄 Конвертация: {os.path.basename(file)}")

                result = self.convert_document(str(file), macro_name)

                if verbose and result:
                    if result.get("status") == "SUCCESS":
                        odt_name = result.get('odt_file', result.get('odt_path', ''))
                        if odt_name:
                            print(f"   ✅ Успешно: {os.path.basename(odt_name)}")
                        else:
                            print(f"   ✅ Успешно сконвертирован")
                        if result.get("deleted"):
                            print(f"   🗑️ Исходный файл удалён")
                    elif result.get("status") == "TIMEOUT":
                        print(f"   ⚠️ Таймаут: макрос не вернул результат")
                        print(f"      Убедитесь, что макрос 'ConvertToODTAndReport' сохранён в Normal.dotm")
                    else:
                        error_msg = result.get('error', 'Неизвестная ошибка')
                        print(f"   ❌ Ошибка: {error_msg}")
        finally:
            self._disconnect_word()

        if verbose:
            self._print_summary()

        return self.results

    def _print_summary(self):
        """Выводит сводку результатов."""
        print("\n" + "=" * 60)
        print("📊 СВОДКА КОНВЕРТАЦИИ WORD → ODT")
        print("=" * 60)

        success = [r for r in self.results if r.get("status") == "SUCCESS"]
        errors = [r for r in self.results if r.get("status") in ["ERROR", "TIMEOUT"]]

        print(f"✅ Успешно сконвертировано: {len(success)} файлов")
        print(f"❌ Ошибок: {len(errors)} файлов")

        if success and config.VERBOSE_OUTPUT:
            print("\n📄 Список сконвертированных файлов:")
            for r in success[:10]:
                odt_name = r.get('odt_file', r.get('odt_path', ''))
                if odt_name:
                    print(f"   - {r.get('original_name')} → {os.path.basename(odt_name)}")
                else:
                    print(f"   - {r.get('original_name')} → ODT")
            if len(success) > 10:
                print(f"   ... и ещё {len(success) - 10} файлов")

        if errors and config.VERBOSE_OUTPUT:
            print("\n❌ Ошибки:")
            for r in errors:
                print(f"   - {r.get('original_name')}: {r.get('error', 'Неизвестная ошибка')}")

    def cleanup_temp_folder(self, confirm=None):
        """Удаляет временную папку с JSON-файлами."""
        if not os.path.exists(self.temp_folder):
            return

        files = glob.glob(os.path.join(self.temp_folder, "*.json"))

        if files and confirm is None:
            print(f"\n📁 Временная папка содержит {len(files)} JSON-файлов")
            choice = input("Удалить временную папку? (y/n): ").strip().lower()
            if choice not in ['y', 'yes', 'д', 'да']:
                print(f"ℹ️ Временная папка сохранена: {self.temp_folder}")
                return

        try:
            shutil.rmtree(self.temp_folder)
            print("✅ Временная папка удалена")
        except Exception as e:
            print(f"⚠️ Не удалось удалить временную папку: {e}")

    def get_stats(self):
        """Возвращает статистику конвертации."""
        return {
            "processed": self.processed_count,
            "success": self.success_count,
            "errors": self.error_count,
            "results": self.results
        }


# ===== ВНЕШНИЕ ФУНКЦИИ ДЛЯ ИНТЕГРАЦИИ С main.py =====

def convert_word_documents_in_folder(folder_path):
    """
    Внешняя функция для вызова из main.py.
    Конвертирует все Word-документы в папке в ODT.
    """
    if not config.CONVERT_WORD_TO_ODT:
        return {"processed": 0, "success": 0, "errors": 0, "results": []}

    converter = WordToODTConverter()
    converter.convert_folder(folder_path)

    # Очищаем временную папку
    converter.cleanup_temp_folder(confirm=False)

    return converter.get_stats()


def needs_word_conversion(filename):
    """Проверяет, нужно ли конвертировать файл Word."""
    ext = os.path.splitext(filename)[1].lower()
    return ext in {'.doc', '.docx'}


def get_word_conversion_stats(converted_files=None):
    """Возвращает статистику по конвертации Word-документов."""
    return 0, 0


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        folder = sys.argv[1]
    else:
        folder = input("Введите путь к папке: ").strip()

    if os.path.isdir(folder):
        convert_word_documents_in_folder(folder)
    else:
        print(f"❌ Папка не найдена: {folder}")