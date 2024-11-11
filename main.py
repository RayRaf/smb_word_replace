from smb.SMBConnection import SMBConnection
from docx import Document
import io
import os

# Параметры подключения
server_name = 'openmediavault.local'
share_name = 'Install'
username = '***'
password = '*******'
file_path = '1.docx'

# Установить соединение
conn = SMBConnection(username, password, 'client_machine_name', server_name, use_ntlm_v2=True)
assert conn.connect(server_name, 139)

# Прочитать файл
file_data = io.BytesIO()
conn.retrieveFile(share_name, file_path, file_data)
file_data.seek(0)

# Работа с файлом Word
doc = Document(file_data)

# Поиск и замена текста
for paragraph in doc.paragraphs:
    if 'старый текст' in paragraph.text:
        paragraph.text = paragraph.text.replace('старый текст', 'новый текст')

# Сохранить изменения в новый файл
new_file_data = io.BytesIO()
doc.save(new_file_data)
new_file_data.seek(0)

# Путь для сохранения нового файла на локальной машине
local_new_file_path = 'documents/example_updated.docx'
os.makedirs(os.path.dirname(local_new_file_path), exist_ok=True)

# Записать данные в новый файл на локальной машине
with open(local_new_file_path, 'wb') as new_file:
    new_file.write(new_file_data.getvalue())

# Загрузить новый файл на сервер
with open(local_new_file_path, 'rb') as new_file:
    conn.storeFile(share_name, file_path, new_file)

# Закрыть соединение
conn.close()
