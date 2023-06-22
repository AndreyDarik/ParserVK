import vk_api  # Импорт модуля vk_api для работы с API ВКонтакте
import tkinter as tk  # Импорт модуля tkinter для создания графического интерфейса
from tkinter import ttk  # Импорт модуля ttk из tkinter для использования темы оформления
from PIL import Image, ImageTk  # Импорт модулей Image и ImageTk из PIL для работы с изображениями
import requests  # Импорт модуля requests для отправки HTTP-запросов
from io import BytesIO  # Импорт модуля BytesIO из io для работы с байтами в памяти
from ttkbootstrap import Style  # Импорт класса Style из ttkbootstrap для создания стиля оформления
import webbrowser
import openpyxl


def search_user():
    global city_var  # Объявление переменных как глобальных

    name = entry.get()  # Получение значения из поля ввода ФИО
    if name:  # Если имя не пустое
        output.delete(1.0, tk.END)  # Очистка поля вывода
        vk_session = vk_api.VkApi(
            token='YOUR_API_TOKEN')  # Создание экземпляра vk_api.VkApi с использованием токена доступа
        vk = vk_session.get_api()  # Получение экземпляра API ВКонтакте

        # Выполнение поиска пользователей в ВКонтакте
        response = vk.users.search(q=name, fields='screen_name, city, country, photo_max_orig', count=30)

        if response['count'] > 0:  # Если найдены пользователи
            count = response['count']
            output.insert(tk.END, f'Найдено пользователей: {count}\n', 'bold')
            for user in response['items']:  # Перебор найденных пользователей
                # Если выбран город и он не совпадает с городом пользователя, пропустить пользователя
                if city_var.get() and user.get('city', {}).get('title') != city_var.get():
                    continue

                output.tag_configure('bold', font=('Arial', 10, 'bold'))  # Конфигурация 'bold' для полужирного шрифта
                output.tag_configure('link', foreground='blue', underline=True)  # Конфигурация 'link' для гиперссылки

                output.insert(tk.END, f'ID: {user["id"]}\n', 'bold')  # Вставка ID пользователя с применением 'bold'
                output.insert(tk.END, f'Имя: {user["first_name"]}\n')  # Вставка имени пользователя
                output.insert(tk.END, f'Фамилия: {user["last_name"]}\n')  # Вставка фамилии пользователя
                output.insert(tk.END, f'Ссылка на профиль: ')  # Вставка текста "Ссылка на профиль: "
                output.insert(tk.END, f'https://vk.com/{user["screen_name"]}\n', 'link')  # Вставка ссылки на профиль
                output.tag_bind('link', '<Button-1>', lambda event, link=user["screen_name"]: open_profile(link))
                output.insert(tk.END, f'Город: {user.get("city", {}).get("title")}\n')  # Вставка названия города
                output.insert(tk.END, f'Страна: {user.get("country", {}).get("title")}\n')  # Вставка названия страны

                photo_url = user.get('photo_max_orig')  # Получение URL фото пользователя
                if photo_url:  # Если URL фото доступен
                    response = requests.get(photo_url)  # Отправка HTTP-запроса для получения фото
                    image = Image.open(BytesIO(response.content))  # Создание изображения из полученных данных
                    image.thumbnail((100, 100))  # Создание миниатюры изображения
                    photo = ImageTk.PhotoImage(image)  # Создание объекта PhotoImage для изображения в tkinter
                    photo_label = ttk.Label(output, image=photo)  # Создание метки с изображением
                    photo_label.image = photo  # Присвоение атрибута image для избежания удаления изображения
                    output.window_create(tk.END, window=photo_label)  # Вставка метки с изображением в поле вывода
                    output.insert(tk.END, '\n')  # Вставка пустой строки
                else:  # Если URL фото не доступен
                    output.insert(tk.END, 'Фото недоступно\n')  # Вставка текста "Фото недоступно"

                output.insert(tk.END, '-------------------------------------\n')  # Вставка разделительной строки
        else:  # Если пользователи не найдены
            output.insert(tk.END, 'Пользователь не найден\n')  # Вставка сообщения "Пользователь не найден"


def handle_enter(event):
    if event.keycode == 13:  # Код клавиши Enter
        search_user()


def open_profile(screen_name):
    url = f'https://vk.com/{screen_name}'
    webbrowser.open(url)


def apply_filters():
    global city_var  # Объявление переменных как глобальных

    city = city_var.get()  # Получение выбранного города из переменной city_var
    output.delete(1.0, tk.END)  # Очистка поля вывода

    # Здесь можно применить фильтры к поиску пользователей в VK
    # Исходный код для применения фильтров

    # Выводим результаты фильтрации
    output.insert(tk.END, 'Применены фильтры:\n')  # Вставка заголовка "Применены фильтры:"
    output.insert(tk.END, f'Город: {city}\n')  # Вставка выбранного города


def export_to_excel():
    global city_var

    name = entry.get()
    if name:  # Если имя не пустое
        output.delete(1.0, tk.END)  # Очистка поля вывода
        # Создание экземпляра vk_api.VkApi с использованием токена доступа
        vk_session = vk_api.VkApi(token='YOUR_API_TOKEN')
        vk = vk_session.get_api()  # Получение экземпляра API ВКонтакте

        response = vk.users.search(q=name, fields='screen_name, city, country, photo_max_orig',
                                   count=100)  # Выполнение поиска пользователей в ВКонтакте

        workbook = openpyxl.Workbook()  # Создание нового документа Excel
        sheet = workbook.active

        sheet.append(['ID', 'Имя', 'Фамилия', 'Ссылка на профиль', 'Город', 'Страна'])  # Запись заголовков столбцов

        for user in response['items']:
            if city_var.get() and user.get('city', {}).get('title') != city_var.get():
                continue

            user_id = user["id"]
            first_name = user["first_name"]
            last_name = user["last_name"]
            screen_name = f'https://vk.com/{user["screen_name"]}'
            city = user.get("city", {}).get("title")
            country = user.get("country", {}).get("title")

            sheet.append([user_id, first_name, last_name, screen_name, city, country])

        filename = 'search_results.xlsx'  # Имя файла для сохранения результатов
        workbook.save(filename)  # Сохранение документа Excel

        output.insert(tk.END, f'Результаты сохранены в файл {filename}\n')


# Создание графического интерфейса
window = tk.Tk()  # Создание окна
window.title('Поиск пользователя VK')  # Установка заголовка окна
window.geometry('1024x960')  # Установка размера окна

# Загрузка изображения фона
background_image = Image.open('background.jpg')  # Замените 'background.jpg' на путь к изображению фона

# Масштабирование изображения фона
background_image = background_image.resize((1024, 960), Image.LANCZOS)

# Преобразование изображения фона в формат Tkinter
background_photo = ImageTk.PhotoImage(background_image)

# Создание метки с фоновым изображением
background_label = ttk.Label(window, image=background_photo)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

# Создание стиля ttkbootstrap
style = Style(theme='journal')

# Загрузка изображений
search_image = Image.open('search_button.png')  # Заменить 'search_button.png' на путь к изображению высокого качества
export_image = Image.open('export_button.png')  # Заменить 'export_button.png' на путь к изображению высокого качества

# Масштабирование изображений
search_image = search_image.resize((32, 32), Image.LANCZOS)
export_image = export_image.resize((32, 32), Image.LANCZOS)

# Преобразование изображений в формат Tkinter
search_button_icon = ImageTk.PhotoImage(search_image)
export_button_icon = ImageTk.PhotoImage(export_image)


frame = ttk.Frame(window)  # Создание фрейма
frame.pack(pady=20)  # Размещение фрейма в окне

label = ttk.Label(frame, text='Введите ФИО:', style='TLabel')  # Создание метки для ввода ФИО
label.grid(row=0, column=0, padx=10, pady=5)  # Размещение метки на фрейме

entry = ttk.Entry(frame, width=30)  # Создание поля ввода ФИО
entry.grid(row=0, column=1, padx=10, pady=5)  # Размещение поля ввода на фрейме

entry.bind('<KeyPress>', handle_enter)  # Привязываем обработчик к событию нажатия клавиши

button = ttk.Button(frame, text='Поиск', image=search_button_icon, compound=tk.LEFT, command=search_user,
                    style='Primary.TButton')  # Создание кнопки поиска
button.grid(row=0, column=2, padx=10, pady=5)  # Размещение кнопки на фрейме

# Создание фрейма для фильтров
filter_frame = ttk.LabelFrame(window, text='Фильтры')
filter_frame.pack(pady=10)

# Создание переменных для фильтров
city_var = tk.StringVar()

# Создание выпадающего списка для выбора города
city_label = ttk.Label(filter_frame, text='Город:')
city_label.grid(row=0, column=0, padx=5, pady=5)

city_combobox = ttk.Combobox(filter_frame, textvariable=city_var, state='readonly')
city_combobox['values'] = ('Москва', 'Санкт-Петербург', 'Сочи', 'Владивосток', 'Казань', 'Екатеринбург',
                           'Нижний Новгород', 'Челябинск')
city_combobox.grid(row=0, column=1, padx=5, pady=5)


# Кнопка применения фильтров
apply_button = ttk.Button(filter_frame, text='Применить фильтры', command=apply_filters)
apply_button.grid(row=2, column=0, columnspan=3, padx=5, pady=5)

# Кнопка экспорта в файл
export_button = ttk.Button(frame, text='Экспорт в Excel', image=export_button_icon, compound=tk.LEFT,
                           command=export_to_excel, style='Primary.TButton')
export_button.grid(row=0, column=3, padx=10, pady=5)

output = tk.Text(window, height=50, width=120, font=('Arial', 10))  # Создание текстового поля вывода
output.pack(pady=20)  # Размещение текстового поля вывода в окне

window.mainloop()  # Запуск главного цикла окна
