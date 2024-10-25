import os
import shutil
import pyautogui
import pygetwindow as gw
import pyperclip
import ctypes
import time
from time import sleep
import tkinter as tk
from tkinter import messagebox

# пауза и досрочное прекращение
pyautogui.PAUSE = 2
pyautogui.FAILSAFE = True

def is_numlock_on():
    if ctypes.WinDLL("User32.dll").GetKeyState(0x90):
        pyautogui.press('numlock') # если numlock включен: отключить 
is_numlock_on()

login = os.getlogin() 
src_dir_container = os.path.join(r'C:\Users', login, 'AppData', 'Local' , 'Infotecs', 'Containers')
pyperclip.copy(src_dir_container)
print(src_dir_container)
               
def show_message():
    # Поднимаем главное окно на передний пла
    root.lift()
    root.attributes('-topmost', True)  # Указываем, что окно должно быть сверху остальных окон
    # Создаем окно сообщения
    messagebox.showinfo("Сообщение", "Сейчас запустится процедура выбора вашего нового сертификата в ViPNet [Деловая почта]\n\nНажмите ОК и не трогайте мышку до появления сообщения")
    # Закрываем программу после нажатия OK
    root.destroy()

# Создаем главное окно
root = tk.Tk()
root.withdraw()  # Скрываем главное окно

# Вызываем функцию для отображения сообщения
show_message()

# Запускаем главный цикл приложения
root.mainloop()

if os.path.isdir(r"C:\Program Files (x86)\InfoTeCS\ViPNet Client"):
    os.startfile(r"C:\Program Files (x86)\InfoTeCS\ViPNet Client\wmail.exe")
    # print("\nДеловая почта запускается. Не трогайте мышку до появления окна с сообщением")
elif os.path.isdir(r"C:\Program Files\Infotecs\ViPNet Client"):
    os.startfile(r"C:\Program Files\Infotecs\ViPNet Client\wmail.exe")
    # print("\nДеловая почта запускается. Не трогайте мышку до появления окна с сообщением")
else:
    # input("У вас не установлена Деловая почта. Нажмите Enter для выхода")
    exit()

# Функция для ожидания появления окна
target_window_title = ("ViPNet Client [Деловая почта]")
def wait_for_window(title, timeout=120):
    start_time = time.time()
    while True:
        all_windows = gw.getAllTitles()
        if title in all_windows:
            # print(f"Окно с заголовком '{title}' появилось!")
            sleep(2)
            break
        if time.time() - start_time > timeout:
            print(f"Время ожидания истекло. Окно '{title}' не появилось.")
            break
        # sleep(1)

# Функция для перемещения окна на основной монитор
def move_window_to_main_monitor(window_title):
    windows = gw.getWindowsWithTitle(window_title)
    if windows:
        window = windows[0]
        main_monitor_position = (0, 0)
        window.moveTo(*main_monitor_position)
        # print(f"Окно '{window_title}' перемещено на основной монитор.")
        return window # Возвращаем объект окна для дальнейших действий
    else:
        print(f"Окно с заголовком '{window_title}' не найдено.")
        return None

# Функция для активации окна
def activate_window(window):
    if window is not None:
        window.activate()

# Основная логика
# start_application_if_exists(app_path)
wait_for_window(target_window_title)
moved_window = move_window_to_main_monitor(target_window_title)
activate_window(moved_window)

# sleep(1)
#инструменты               
pyautogui.press('f10')
pyautogui.press('right', presses=2, interval=0.2) 
pyautogui.press('enter') 

# sleep(1)
#Настройка параметров безопасности>Ключи
pyautogui.press('down', presses=4, interval=0.2)
pyautogui.press('enter')
pyautogui.press('tab', presses=7, interval=0.2)
pyautogui.press('up')
pyautogui.press('right')
pyautogui.press('tab', presses=4, interval=0.2)
pyautogui.press('enter')

# sleep(1)
#инициализация контейнера
pyautogui.press('tab')
pyautogui.hotkey('shift', 'insert')

pyautogui.press('enter')

# sleep(2)

pyautogui.press('up')
pyautogui.press('enter')

# Функция для отображения сообщения
def show_message():

    # Создаем окно сообщения
    root.lift()
    root.attributes('-topmost', True)  # Указываем, что окно должно быть сверху остальных окон
    messagebox.showinfo("Сообщение", "Новый сертификат выбран\nВ окне 'Настройка параметров безопасности' нажмите ОК (окно может зависнуть)")

    root.destroy()  # Закрываем главное окно после нажатия OK

# Создаем главное окно
root = tk.Tk()
root.withdraw()  # Скрываем главное окно

# Вызываем функцию для отображения сообщения
show_message()

# Запускаем главный цикл приложения
root.mainloop()


