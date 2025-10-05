# Используем официальный образ Python
FROM python:3.11-slim

# Устанавливаем рабочую директорию
WORKDIR /app

# Копируем файлы проекта
COPY . /app

# Устанавливаем зависимости
RUN pip install --no-cache-dir aiogram

# Если потребуется, добавьте другие зависимости:
# RUN pip install --no-cache-dir <package>

# Указываем переменную окружения для корректной работы Python
ENV PYTHONUNBUFFERED=1

# Запускаем бота
CMD ["python", "main.py"]
