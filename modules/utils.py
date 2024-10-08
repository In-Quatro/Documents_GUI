import os


def get_new_file_name(file_name, folder):
    """Переименовывает новый файл если уже есть такой файл в папке."""
    base, ext = os.path.splitext(file_name)
    index = 1
    new_file_name = file_name

    while os.path.exists(os.path.join(folder, new_file_name)):
        new_file_name = f'{base} ({index}){ext}'
        index += 1
    return new_file_name
