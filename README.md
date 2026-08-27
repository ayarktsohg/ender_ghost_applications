# Приложение для сравнения файлов с конвертацией

Консольное приложение для сравнения двух папок с рекурсивным обходом и автоматической конвертацией изображений и документов Word.

## Возможности

- Рекурсивное сравнение папок и файлов (по названию и размеру)
- Конвертация изображений (JPG, BMP, GIF, WEBP) в PNG
- Конвертация документов Word (DOC, DOCX) в ODT
- Подробный отчёт о результатах сравнения и конвертации

## Требования

- **Python 3.9+**
- **Microsoft Word** (для конвертации DOC/DOCX в ODT) — требуется установленная лицензионная копия
- **LibreOffice** (опционально, для альтернативной конвертации)

## Установка

1. Клонируйте репозиторий:
```bash
git clone https://github.com/ayarktsohg/ender_ghost_applications.git
cd folder-comparator
```
2. Создайте виртуальное окружение:
```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
# или
source .venv/bin/activate  # Linux/Mac
```
3. Установите зависимости:
```bash
pip install -r requirements.txt
```
## Использование
```bash
python main.py
```
Следуйте инструкциям в консоли:

1. Введите путь к первой папке
2. Введите путь ко второй папке
3. Программа выполнит конвертацию (если включена) и сравнение

## Настройка
Все настройки находятся в файле config.py:
- ```CONVERT_IMAGES_TO_PNG``` — включить конвертацию изображений
- ```CONVERT_WORD_TO_ODT``` — включить конвертацию Word → ODT
- ```DELETE_ORIGINAL_AFTER_CONVERT``` — удалять ли исходные файлы после конвертации
## Требования к макросу Word
Для конвертации DOC/DOCX в ODT необходимо добавить макрос ConvertToODTAndReport в шаблон Normal.dotm.
## Лицензия
Этот проект распространяется под лицензией MIT. Подробнее в файле LICENSE.

## Зависимости от стороннего ПО

- **Microsoft Word** — для конвертации документов. Требуется установленная лицензионная копия Microsoft Office.
- **Pillow** — лицензия PIL/MIT
- **pywin32** — лицензия Python Software Foundation


---


## Установка макроса для конвертации Word → ODT

Для работы конвертации документов Word в ODT необходимо добавить макрос в шаблон `Normal.dotm`.

## Инструкция

1. Откройте **Microsoft Word**
2. Нажмите `Alt + F11` (откроется редактор VBA)
3. В окне слева (Project Explorer) найдите **Normal** → **Modules**
4. Если модуля нет, создайте: **Insert → Module**
5. Скопируйте и вставьте код макроса (см. ниже)
6. Нажмите `Ctrl + S` для сохранения
7. Закройте редактор VBA (`Alt + Q`)

## Код макроса

```vb
Sub ConvertToODTAndReport()
    ' Макрос конвертирует активный документ в ODT
    ' и записывает JSON-отчёт во временную папку (UTF-8)
    
    On Error GoTo ErrorHandler
    
    Dim docName As String
    Dim docPath As String
    Dim odtPath As String
    Dim tempFolder As String
    Dim resultFile As String
    Dim jsonContent As String
    Dim fsoCheck As Object
    Dim stream As Object
    
    ' 1. Получаем информацию о документе
    With ActiveDocument
        docName = .Name
        docPath = .FullName
    End With
    
    ' 2. Формируем путь к ODT-файлу (в той же папке)
    odtPath = Left(docPath, InStrRev(docPath, ".") - 1) & ".odt"
    
    ' 3. Проверяем, не существует ли уже ODT-файл
    Set fsoCheck = CreateObject("Scripting.FileSystemObject")
    If fsoCheck.FileExists(odtPath) Then
        Dim counter As Integer
        counter = 1
        Do While fsoCheck.FileExists(odtPath)
            odtPath = Left(docPath, InStrRev(docPath, ".") - 1) & "_" & counter & ".odt"
            counter = counter + 1
        Loop
    End If
    
    ' 4. Сохраняем как ODT
    ActiveDocument.SaveAs2 FileName:=odtPath, FileFormat:=wdFormatOpenDocumentText
    
    ' 5. Формируем JSON-отчёт (экранируем кавычки)
    jsonContent = "{" & vbCrLf & _
                  "  ""status"": ""SUCCESS""," & vbCrLf & _
                  "  ""original_file"": """ & Replace(docName, """", "\""") & """," & vbCrLf & _
                  "  ""odt_file"": """ & Mid(odtPath, InStrRev(odtPath, "\") + 1) & """," & vbCrLf & _
                  "  ""original_path"": """ & Replace(docPath, "\", "\\") & """," & vbCrLf & _
                  "  ""odt_path"": """ & Replace(odtPath, "\", "\\") & """," & vbCrLf & _
                  "  ""timestamp"": """ & Now & """" & vbCrLf & _
                  "}"
    
    ' 6. Создаём временную папку
    tempFolder = Environ("TEMP") & "\word_converter_temp"
    If Not fsoCheck.FolderExists(tempFolder) Then
        fsoCheck.CreateFolder tempFolder
    End If
    
    ' 7. Создаём уникальное имя файла
    Dim uniqueId As String
    uniqueId = Format(Now, "yyyymmddhhmmss") & "_" & Int(Rnd * 100000)
    resultFile = tempFolder & "\result_" & uniqueId & ".json"
    
    ' 8. Записываем JSON в UTF-8 (без BOM)
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText jsonContent
    stream.SaveToFile resultFile, 2 ' adSaveCreateOverWrite
    stream.Close
    
    ' 9. Сохраняем имя файла в переменную среды
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.Environment("User")("LAST_CONVERT_RESULT") = resultFile
    
    Exit Sub
    
ErrorHandler:
    ' В случае ошибки записываем JSON с ошибкой
    Dim errContent As String
    errContent = "{" & vbCrLf & _
                 "  ""status"": ""ERROR""," & vbCrLf & _
                 "  ""original_file"": """ & Replace(ActiveDocument.Name, """", "\""") & """," & vbCrLf & _
                 "  ""error"": """ & Replace(Err.Description, """", "\""") & """," & vbCrLf & _
                 "  ""error_number"": " & Err.Number & "," & vbCrLf & _
                 "  ""timestamp"": """ & Now & """" & vbCrLf & _
                 "}"
    
    tempFolder = Environ("TEMP") & "\word_converter_temp"
    Set fsoCheck = CreateObject("Scripting.FileSystemObject")
    If Not fsoCheck.FolderExists(tempFolder) Then
        fsoCheck.CreateFolder tempFolder
    End If
    
    Dim errorFile As String
    errorFile = tempFolder & "\error_" & Format(Now, "yyyymmddhhmmss") & ".json"
    
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText errContent
    stream.SaveToFile errorFile, 2
    stream.Close
    
    Dim shellErr As Object
    Set shellErr = CreateObject("WScript.Shell")
    shellErr.Environment("User")("LAST_CONVERT_RESULT") = errorFile
End Sub
```

## Проверка

Чтобы проверить, что макрос работает:

1. Откройте любой .docx файл в Word 
2. Нажмите Alt + F8
3. Выберите ConvertToODTAndReport
4. Нажмите "Выполнить"
5. Проверьте папку %TEMP%\word_converter_temp — там должен появиться JSON-файл


---

