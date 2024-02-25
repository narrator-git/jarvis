import pvporcupine
import os
import random
import pyaudio
import pygame
import ctypes
import numpy as np
import speech_recognition as sr
import webbrowser
import json
import pyautogui
import time
import pygetwindow as gw
import pyperclip
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import ctypes

#Глобальные переменные
greetings = 'C:\\Users\\mrayd\\Desktop\\my programms\\Приветствия'
doing = 'C:\\Users\\mrayd\\Desktop\\my programms\\Выполнение'
fishing = 'C:\\Users\\mrayd\\Desktop\\my programms\\Рыбалка'
end = 'C:\\Users\\mrayd\\Desktop\\my programms\\Конец'
websites_list = [
    "https://www.youtube.com/","https://chat.openai.com/c/88d577b7-1a26-4326-80d2-bd0084b3ec17",
    "https://web.whatsapp.com/","https://www.instagram.com/","https://mail.google.com/mail/u/0/#inbox"
]



#Комманды
commands_list=[
    #вебсайты
    "ютьюб","нейросеть","ватсап","инстаграм","почту",
    #открой приложения
    "открой споти","открой телеграм","открой гугл",
    #закрой
    "закрой окно","закрой вкладку",
    #другое
    "покажи на тв","останови показ","перезагрузись"]






def find_google_window():
    for window_title in gw.getAllTitles():
        if "Google" in window_title:
            return window_title
    return None

def activate_google_window():
    google_window = find_google_window()
    if google_window:
        window = gw.getWindowsWithTitle(google_window)[0]
        window.activate()


def get_keyboard_layout():
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    thread_id = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, 0)
    klid = ctypes.windll.user32.GetKeyboardLayout(thread_id)
    layout_id = klid & 0xFFFF
    return hex(layout_id)








# Функция для обработки команды после распознавания ключевого слова
def process_command(command):
    # Здесь вы можете добавить код для обработки команды и генерации ответа от системы
    # Например, вы можете включить другую функцию, которая будет отвечать на команду "Как дела?"
    if "откр" in command:
        comm=""
        words = command.split()
        try:
            index = words.index("открыть")
        except:
            index = words.index("открой")
        if index + 1 < len(words):
            comm = words[index + 1:index+2]
        max_percentage=0
        for i in range (5):
            percentage = similarity_percentage(str(comm), commands_list[i])
            if percentage>max_percentage and percentage>25:
                max_percentage=percentage
                command=i
            play_random_mp3(doing)
        try:
            webbrowser.open(websites_list[command])
        except:
            webbrowser.open("https://duckduckgo.com/?q="+" ".join(comm)+"&ia=web")
    elif "поиск" in command:
        comm=""
        words = command.split()
        index = words.index("поиск")
        if index + 1 < len(words):
            comm = words[index + 1:]       
            webbrowser.open("https://duckduckgo.com/?q="+" ".join(comm)+"&ia=web")
            play_random_mp3(doing)

    elif "пиши" in command:
        comm=""
        words = command.split()
        try:
            index = words.index("пиши")
        except:
            index = words.index("напиши")
        if index + 1 < len(words):
            comm = words[index + 1::]
        pyperclip.copy(" ".join(comm))




    elif "став" in command:
        pyautogui.hotkey("ctrl","v")




    elif "отправ" in command:
        pyautogui.press("enter")




    elif "включи" in command:
        comm=""
        words = command.split()
        index = words.index("включи")
        if index + 1 < len(words):
            comm = words[index + 1::]
        if len(comm)!=0:
            for i in range(4):
                layout = get_keyboard_layout()
                if layout=="0x409":
                    break
                else:
                    pyautogui.hotkey("alt","shift")
            time.sleep(1)
            pyperclip.copy(" ".join(comm))
            pyautogui.press("win")
            time.sleep(1)  # Подождем секунду, чтобы меню "Пуск" открылось полностью

            # Вводим название приложения в поисковую строку
            pyautogui.write("Spotify", interval=0.1)
            time.sleep(1)  # Даем время для поиска приложения

            # Нажимаем Enter для запуска приложения
            pyautogui.press("enter")
            time.sleep(3)

            pyautogui.moveTo(52,146)
            pyautogui.click()
            time.sleep(1)
            pyautogui.moveTo(346,86)
            pyautogui.click()
            time.sleep(1)
            pyautogui.hotkey("ctrl","v")
            time.sleep(2)
            pyautogui.moveTo(672,519)
            pyautogui.click()
        else:
            pyautogui.press("playpause")

    elif "гром" in command:
        for i in range(5):
            pyautogui.press("volumeup")
            time.sleep(0.5)
    elif "тиш" in command:
        for i in range(5):
            pyautogui.press("volumedown")
            time.sleep(0.5)
    elif "пауз" in command:
        pyautogui.press("playpause")
    elif "клик" in command and "прав" in command:
        pyautogui.rightClick()
    elif "рыба" in command:
        play_random_mp3(fishing)
        for i in range (10):
            pyautogui.rightClick()
            time.sleep(5)
            pyautogui.rightClick()
        play_random_mp3(end)
        


#Алгоритм расстояния Левенштайна
def similarity_percentage(s1, s2):
    len1, len2 = len(s1), len(s2)
    if len1 < len2:
        s1, s2 = s2, s1
        len1, len2 = len2, len1

    previous_row = range(len2 + 1)
    for i, c1 in enumerate(s1, 1):
        current_row = [i]
        for j, c2 in enumerate(s2, 1):
            insertions = current_row[j - 1] + 1
            deletions = previous_row[j] + 1
            substitutions = previous_row[j - 1] + (c1 != c2)
            current_row.append(min(insertions, deletions, substitutions))
        previous_row = current_row

    distance = previous_row[-1]
    max_len = max(len1, len2)
    similarity = 1 - distance / max_len
    return similarity * 100









# Функция для воспроизведения случайного аудиофайла
def play_random_mp3(folder_path):
    mp3_files = [file for file in os.listdir(folder_path) if file.endswith(".wav")]
    random_mp3 = random.choice(mp3_files)
    mp3_file_path = os.path.join(folder_path, random_mp3)
    pygame.mixer.init()
    pygame.mixer.music.load(mp3_file_path)
    pygame.mixer.music.play()













# Функция для распознавания речи с помощью Vosk
def recognize_speech():
    recognizer = sr.Recognizer()
    mic = sr.Microphone()
    with mic as source:
            recognizer.adjust_for_ambient_noise(source)
            print("Скажите команду:")
            audio = recognizer.listen(source)
    try:
        recognized_text = recognizer.recognize_vosk(audio)
        data = json.loads(recognized_text)
        word = data["text"]
        print("Вы сказали:", word)
        return word
    except sr.UnknownValueError:
        print("Извините, не удалось распознать речь.")
        return None
    except sr.RequestError as e:
        print(f"Ошибка при запросе к сервису распознавания речи; {e}")
        return None














# Основная функция для обнаружения ключевого слова и распознавания речи
def is_wake_word(keyword_path):
    handle = pvporcupine.create(keyword_paths=[keyword_path], access_key="eA0famHpKN24sn6LkxHvwyRbsGBlZxsGf4koWrM9dZqigJ1OoEjD8A==")

    audio = pyaudio.PyAudio()
    stream = audio.open(rate=handle.sample_rate, channels=1, format=pyaudio.paInt16, input=True, frames_per_buffer=handle.frame_length)

    while True:
        pcm = stream.read(handle.frame_length)
        pcm = np.frombuffer(pcm, dtype=np.int16)

        keyword_index = handle.process(pcm)

        if keyword_index >= 0:
            print("Ключевое слово 'Jarvis' обнаружено.")
            play_random_mp3(greetings)

            # Запись и распознавание речи после обнаружения ключевого слова
            recognized_text = recognize_speech()
            if "перезагрузись" in recognized_text:
                # Задержка перед выполнением следующего действия
                pyautogui.sleep(1)
                # Нажатие клавиши Win+R для открытия окна "Выполнить"
                pyautogui.hotkey('win', 'r')
                # Задержка перед выполнением следующего действия
                pyautogui.sleep(1)
                # Вставка пути к вашему Python скрипту
                pyautogui.write('C:\\Users\\mrayd\\Desktop\\my programms\\Restart.py')
                # Нажатие клавиши Enter для запуска скрипта
                pyautogui.press('enter')
                pyautogui.sleep(1)
                break
            if recognized_text!="":
                process_command(recognized_text)

    handle.delete()















if __name__ == "__main__":
    keyword_path = 'Jarvis_en_windows_v2_2_0.ppn'
    is_wake_word(keyword_path)
