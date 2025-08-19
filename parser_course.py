from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
import os
from selenium.webdriver.common.keys import Keys

# Глобальный кеш профилей ВК, проверенных в текущем запуске
processed_vk_profiles: set[str] = set()


def open_vk_homework_page():
    options = Options()
    options.add_argument("user-data-dir=G:/SHA_VK/chrome_profile")
    
    # Дополнительные опции для стабильности
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-plugins")
    options.add_argument("--disable-images")
    options.add_argument("--disable-javascript")
    options.add_argument("--disable-web-security")
    options.add_argument("--allow-running-insecure-content")
    
    # Ускоряем загрузку страниц
    try:
        options.page_load_strategy = 'eager'
    except Exception:
        pass
    
    # Отключаем картинки для ускорения
    try:
        options.add_experimental_option("prefs", {
            "profile.managed_default_content_settings.images": 2,
            "profile.managed_default_content_settings.javascript": 1
        })
    except Exception:
        pass
    
    driver = webdriver.Chrome(options=options)
    
    # Увеличиваем таймауты
    try:
        driver.set_page_load_timeout(30)  # уменьшаем с 60 до 30 секунд
        driver.implicitly_wait(3)  # уменьшаем с 10 до 3 секунд
    except Exception:
        pass
    
    print("🌐 Открываем страницу с домашками...")
    driver.get("https://authors.vk.company/profile/v.chernikov/homework/?type=ready&owner=all&p=1")
    return driver


def go_to_last_homework(driver):
    wait = WebDriverWait(driver, 15)  # уменьшаем с 20 до 15 секунд
    # Ждём появления кнопки последней страницы
    last_page_btn = wait.until(EC.element_to_be_clickable((
        By.CSS_SELECTOR,
        "#react-homeworks > div > div.sc-gFqAkR.kiXkom > div > button.r-button.button-pagination.boundary > span"
    )))
    last_page_btn.click()

    # Ждём загрузки таблицы с домашками (уменьшаем таймаут)
    wait.until(EC.presence_of_element_located((
        By.CSS_SELECTOR,
        "#react-homeworks > div > div.sc-gFqAkR.kiXkom > table > tbody > tr"
    )))
    time.sleep(5)  # уменьшаем с 8 до 5 секунд

    # Находим все строки таблицы
    rows = driver.find_elements(By.CSS_SELECTOR, "#react-homeworks > div > div.sc-gFqAkR.kiXkom > table > tbody > tr")
    last_row = rows[-1]
    # В последней строке ищем ссылку на домашку
    last_hw_link = last_row.find_element(By.CSS_SELECTOR, "td.sc-eqUAAy.sc-iGgWBj.jjpiPE.gbwPlL > a")
    
    # Открываем ссылку в новой вкладке (Ctrl+Click)
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", last_hw_link)
    time.sleep(3)  # уменьшаем с 5 до 3 секунд
    
    # Открываем ссылку в новой вкладке через JavaScript
    homework_url = last_hw_link.get_attribute("href")
    driver.execute_script("window.open(arguments[0], '_blank');", homework_url)
    time.sleep(0.5)  # уменьшаем с 1 до 0.5 секунды
    
    # Переключаемся на новую вкладку
    driver.switch_to.window(driver.window_handles[-1])
    
    # Ждём загрузки новой страницы
    wait.until(EC.presence_of_element_located((
        By.CSS_SELECTOR,
        "#content > div.homework-chat-header > div.homework-chat-header-left > h1"
    )))


def process_homework_page(driver):
    wait = WebDriverWait(driver, 5)  # уменьшаем с 8 до 5 секунд
    
    # Быстрая проверка - есть ли уже Владимир Черников
    try:
        tutors_block = driver.find_element(By.CSS_SELECTOR, "#homework-tutors .block-content")
        tutor_links = tutors_block.find_elements(By.CSS_SELECTOR, "a.user-name")
        for link in tutor_links:
            if link.text.strip() == "Владимир Черников":
                print("Владимир Черников уже назначен. Пропускаем.")
                return
    except Exception:
        pass

    # 1. Клик по кнопке назначения
    btn = wait.until(EC.element_to_be_clickable((
        By.CSS_SELECTOR,
        "#homework-tutors .buttons-box button"
    )))
    btn.click()

    # 2. Клик по полю поиска
    search_input = wait.until(EC.element_to_be_clickable((
        By.CSS_SELECTOR,
        "#homework-tutors .block-content input"
    )))
    search_input.click()

    # 3. Найти и отметить Владимира Черникова
    labels = driver.find_elements(By.CSS_SELECTOR, "#homework-tutors .search-label")
    for label in labels:
        if label.text.strip() == "Владимир Черников":
            parent_div = label.find_element(By.XPATH, "..")
            checkbox = parent_div.find_element(By.CSS_SELECTOR, "input[type=checkbox]")
            if not checkbox.is_selected():
                driver.execute_script("arguments[0].click();", checkbox)
                print("✅ Чекбокс отмечен")
            break

    # 4. Закрываем список через ESC (быстрее)
    driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
    time.sleep(0.1)

    # 5. Сохраняем
    confirm_btn = wait.until(EC.element_to_be_clickable((
        By.CSS_SELECTOR,
        "#homework-tutors .buttons-box .btn.btn-primary"
    )))
    driver.execute_script("arguments[0].click();", confirm_btn)
    print("✅ Проверяющий назначен")
    
    # Ждём загрузки диалога
    try:
        wait.until(EC.presence_of_element_located((
            By.CSS_SELECTOR,
            "#react-talk .talk"
        )))
    except Exception:
        pass

def remove_from_reviewers(driver):
    """Удаляет Владимира Черникова из проверяющих, если там больше одного человека"""
    wait = WebDriverWait(driver, 8)  # уменьшаем с 10 до 8 секунд
    
    max_attempts = 3  # максимум 3 попытки
    attempt = 1
    
    while attempt <= max_attempts:
        try:
            # Убираем детальное логирование для ускорения
            
            # Проверяем количество проверяющих
            tutors_block = driver.find_element(By.CSS_SELECTOR, "#homework-tutors .block-content")
            tutor_users = tutors_block.find_elements(By.CSS_SELECTOR, "div.user.user-md")
            
            if len(tutor_users) <= 1:
                print("Владимир Черников единственный - не удаляем")
                return
            
            print(f"Найдено проверяющих: {len(tutor_users)}. Удаляем Владимира Черникова...")
            
            # 1. Клик по кнопке изменения
            change_btn = driver.find_element(By.CSS_SELECTOR, "#homework-tutors .buttons-box button")
            change_btn.click()
            time.sleep(0.5)  # уменьшаем с 1 до 0.5 секунды
            
            # 2. Клик по полю ввода
            input_field = driver.find_element(By.CSS_SELECTOR, "#homework-tutors .block-content input")
            input_field.click()
            time.sleep(0.3)  # уменьшаем с 1 до 0.3 секунды
            
            # 3. Найти и снять чекбокс с Владимира Черникова
            labels = driver.find_elements(By.CSS_SELECTOR, "#homework-tutors .search-label")
            for label in labels:
                if label.text.strip() == "Владимир Черников":
                    parent_div = label.find_element(By.XPATH, "..")
                    checkbox = parent_div.find_element(By.CSS_SELECTOR, "input[type=checkbox]")
                    
                    if checkbox.is_selected():
                        driver.execute_script("arguments[0].click();", checkbox)
                        print("Снят чекбокс с Владимира Черникова")
                    break
            
            # 4. Закрываем список через ESC
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            time.sleep(0.3)  # уменьшаем с 1 до 0.3 секунды
            
            # 5. Сохраняем
            save_btn = driver.find_element(By.CSS_SELECTOR, "#homework-tutors .buttons-box .btn.btn-primary")
            driver.execute_script("arguments[0].click();", save_btn)
            time.sleep(1)  # уменьшаем с 2 до 1 секунды
            print("✅ Владимир Черников удалён из проверяющих")
            return  # Успешно завершаем
            
        except Exception as e:
            print(f"❌ Ошибка при удалении из проверяющих (попытка {attempt}): {e}")
            
            if "stale element" in str(e).lower():
                print(f"🔄 Stale element error - пробуем ещё раз с увеличенными паузами...")
                attempt += 1
                if attempt <= max_attempts:
                    # Увеличиваем паузы при повторных попытках
                    time.sleep(2)  # дополнительная пауза перед повтором
                    continue
            else:
                print(f"❌ Критическая ошибка, не связанная со stale element: {e}")
                break
    
    print(f"❌ Не удалось удалить из проверяющих после {max_attempts} попыток")


def extract_cloud_links(driver):
    # Убираем детальное логирование времени для ускорения
    cloud_links = []
    
    print("🔍 Быстрый поиск ссылок на облако...")
    
    # 1. ПОИСК ССЫЛОК В ДИАЛОГЕ - используем более надёжный подход
    try:
        print("🔍 Ищем ссылки в диалоге...")
        # Убираем детальное логирование времени для ускорения
        
        # Сначала пробуем найти все сообщения пользователей
        try:
            # Ищем все сообщения пользователей
            user_messages = driver.find_elements(By.CSS_SELECTOR, "#react-talk .message-user .text")
            print(f"🔍 Найдено {len(user_messages)} сообщений пользователей")
            
            for i, message in enumerate(user_messages):
                try:
                    # Получаем HTML содержимое сообщения
                    message_html = message.get_attribute("innerHTML")
                    message_text = message.text
                    
                    # Убираем детальное логирование для ускорения
                    
                    # Ищем ссылки в HTML (теги <a>)
                    href_pattern = r'href=["\']([^"\']+)["\']'
                    hrefs_in_html = re.findall(href_pattern, message_html)
                    
                    if hrefs_in_html:
                        # Убираем детальное логирование для ускорения
                        for href in hrefs_in_html:
                            if href and href.startswith('http'):
                                if is_valid_cloud_link(href):
                                    cloud_links.append(href)
                                    print(f"✅ Добавлена ссылка из HTML сообщения {i+1}: {href}")
                                else:
                                    print(f"❌ Пропускаем невалидную ссылку из HTML: {href}")
                    
                    # Ищем URL в тексте сообщения (если HTML не дал результатов)
                    if not hrefs_in_html and message_text:
                        url_pattern = r'https?://[^\s<>"{}|\\^`\[\]]+'
                        urls_in_text = re.findall(url_pattern, message_text)
                        
                        if urls_in_text:
                            # Убираем детальное логирование для ускорения
                            for url in urls_in_text:
                                if is_valid_cloud_link(url):
                                    cloud_links.append(url)
                                    print(f"✅ Добавлена ссылка из текста сообщения {i+1}: {url}")
                                else:
                                    print(f"🔍 Пропускаем невалидную ссылку из текста: {url}")
                    
                except Exception as e:
                    print(f"🔍 Ошибка при обработке сообщения {i+1}: {e}")
                    continue
            
        except Exception as e:
            print(f"🔍 Не удалось найти сообщения пользователей: {e}")
            # Fallback: ищем по всему диалогу
            try:
                dialog_element = driver.find_element(By.CSS_SELECTOR, "#react-talk .talk")
                dialog_html = dialog_element.get_attribute("innerHTML")
                dialog_text = dialog_element.text
                
                # Убираем детальное логирование для ускорения
                
                # Ищем ссылки в HTML диалога
                href_pattern = r'href=["\']([^"\']+)["\']'
                hrefs_in_dialog = re.findall(href_pattern, dialog_html)
                
                if hrefs_in_dialog:
                                            # Убираем детальное логирование для ускорения
                    for href in hrefs_in_dialog:
                        if href and href.startswith('http'):
                            if is_valid_cloud_link(href):
                                cloud_links.append(href)
                                print(f"✅ Добавлена ссылка из HTML диалога: {href}")
                            else:
                                print(f"❌ Пропускаем невалидную ссылку из HTML диалога: {href}")
                
                # Ищем URL в тексте диалога
                if dialog_text:
                    url_pattern = r'https?://[^\s<>"{}|\\^`\[\]]+'
                    urls_in_dialog_text = re.findall(url_pattern, dialog_text)
                    
                    if urls_in_dialog_text:
                        # Убираем детальное логирование для ускорения
                        for url in urls_in_dialog_text:
                            if is_valid_cloud_link(url):
                                cloud_links.append(url)
                                print(f"✅ Добавлена ссылка из текста диалога: {url}")
                            else:
                                print(f"🔍 Пропускаем невалидную ссылку из текста диалога: {url}")
                
            except Exception as e2:
                print(f"🔍 Ошибка при fallback поиске: {e2}")
        
        # Убираем детальное логирование времени для ускорения
        
    except Exception as e:
        print(f"❌ Ошибка поиска в диалоге: {e}")
    
    # 2. ПОИСК ССЫЛОК В КОММЕНТАРИИ
    try:
        print("🔍 Ищем ссылки в комментарии...")
        # Убираем детальное логирование времени для ускорения
        
        # Ищем ссылки в HTML комментария
        comment_div = driver.find_element(By.CSS_SELECTOR, "#homework-panel .content-renderer")
        comment_html = comment_div.get_attribute("innerHTML")
        
        if comment_html:
            # Быстрый regex для всех href
            href_pattern = r'href=["\']([^"\']+)["\']'
            hrefs = re.findall(href_pattern, comment_html)
            
            for href in hrefs:
                if href and href.startswith('http'):
                    # Используем улучшенную функцию фильтрации
                    if is_valid_cloud_link(href):
                        cloud_links.append(href)
                        print(f"✅ Добавлена ссылка на облако из комментария: {href}")
                    else:
                        print(f"❌ Пропускаем невалидную ссылку из комментария: {href}")
                else:
                    print(f"🔍 Пропускаем не-HTTP ссылку из комментария: {href}")
            
            # Если HTML-поиск не дал результатов, ищем по тексту комментария
            if not hrefs:
                try:
                    print("🔍 HTML-поиск в комментарии не дал результатов, ищем по тексту...")
                    comment_text = comment_div.text
                    if comment_text:
                        # Ищем URL в тексте
                        url_pattern = r'https?://[^\s<>"{}|\\^`\[\]]+'
                        urls_in_text = re.findall(url_pattern, comment_text)
                        # Убираем детальное логирование для ускорения
                        
                        for url in urls_in_text:
                            if is_valid_cloud_link(url):
                                cloud_links.append(url)
                                print(f"🔍 Найдена ссылка в тексте комментария: {url}")
                            else:
                                print(f"🔍 Пропускаем невалидную ссылку в тексте комментария: {url}")
                except Exception as e:
                    print(f"🔍 Ошибка при поиске по тексту комментария: {e}")
        
        # Убираем детальное логирование времени для ускорения
                    
    except Exception as e:
        print(f"❌ Ошибка поиска в комментарии: {e}")
    
    # Убираем дубликаты
    cloud_links = list(set(cloud_links))
    
    # Убираем детальное логирование времени для ускорения
    if cloud_links:
        print(f"✅ Найдены ссылки на облако ({len(cloud_links)} шт.)")
        for link in cloud_links:
            print(f"   • {link}")
    else:
        print(f"❌ Ссылки на облако не найдены")
    
    return cloud_links


def create_excel_table():
    # Создаём DataFrame с нужными столбцами
    columns = [
        'КО', 'Группа', 'ФИ студента с платформы', 'ВХ Анкета (У1)',
        'Ссылка на ДЗ №1', 'Комментарий к ДЗ №1', 'Оценка по ДЗ №1',
        'Ссылка на ДЗ №2', 'Комментарий к ДЗ №2', 'Оценка по ДЗ №2',
        'Ссылка на профиль на платформе', 'Ссылка на страницу ВКонтакте',
        'ФИ из ВКонтакте (если отличается)', 'Ссылка на сообщество',
        'Количество подписчиков', 'Сумма'
    ]
    
    filename = "homework_data.xlsx"
    
    # Проверяем, существует ли файл
    if os.path.exists(filename):
        print(f"Файл {filename} уже существует, будем дополнять данные")
        return
    
    # Создаём новый файл только если его нет
    df = pd.DataFrame(columns=columns)
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Данные')
        
        # Получаем рабочий лист для настройки ширины колонок
        worksheet = writer.sheets['Данные']
        
        # Настраиваем ширину колонок под содержимое
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            # Устанавливаем ширину колонки (минимальная ширина 15, максимальная 50)
            adjusted_width = min(max(max_length + 2, 15), 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f"Создана новая таблица: {filename}")


def extract_student_data(driver):
    wait = WebDriverWait(driver, 5)  # уменьшаем с 8 до 5 секунд
    student_data = {}
    
    # 1. Извлекаем ФИ пользователя и ссылку на профиль
    try:
        user_name_element = driver.find_element(By.CSS_SELECTOR, "#homework-marks .user-name")
        student_data['full_name'] = user_name_element.text.strip()
        student_data['profile_link'] = user_name_element.get_attribute("href")
        print(f"ФИ студента: {student_data['full_name']}")
        print(f"Ссылка на профиль: {student_data['profile_link']}")
    except Exception as e:
        print(f"Ошибка при получении ФИ студента: {e}")
        student_data['full_name'] = ""
        student_data['profile_link'] = ""
    
    # 2. Извлекаем количество баллов
    try:
        mark_element = driver.find_element(By.CSS_SELECTOR, "#homework-marks .mark-value")
        mark_text = mark_element.text.strip()
        # Извлекаем только цифру из текста "16 баллов"
        mark_match = re.search(r'(\d+)', mark_text)
        if mark_match:
            student_data['mark'] = int(mark_match.group(1))
            print(f"Количество баллов: {student_data['mark']}")
        else:
            student_data['mark'] = ""
            print("Не удалось извлечь количество баллов")
    except Exception as e:
        print(f"Ошибка при получении баллов: {e}")
        student_data['mark'] = ""
    
    return student_data

def extract_profile_data(driver):
    wait = WebDriverWait(driver, 5)  # уменьшаем с 8 до 5 секунд
    profile_data = {}
    
    # 1. Извлекаем номер группы
    try:
        group_element = driver.find_element(By.CSS_SELECTOR, "#profile-content .profile-maingroup")
        # Ищем все ссылки внутри элемента
        group_links = group_element.find_elements(By.TAG_NAME, "a")
        group_number = ""
        for link in group_links:
            text = link.text.strip()
            # Проверяем, что группа начинается с ТБ
            if text.startswith("ТБ"):
                group_number = text
                break
        
        profile_data['group'] = group_number
        if group_number:
            print(f"Номер группы: {group_number}")
        else:
            print("Группа, начинающаяся с ТБ, не найдена")
    except Exception as e:
        print(f"Ошибка при получении номера группы: {e}")
        profile_data['group'] = ""
    
    # 2. Извлекаем ссылку на ВКонтакте
    try:
        vk_element = driver.find_element(By.CSS_SELECTOR, "#content .profile-right .profile-external-accounts a")
        vk_url = vk_element.get_attribute("href")
        # Проверяем, что это действительно ссылка на vk.com
        if vk_url and "vk.com" in vk_url:
            profile_data['vk_link'] = vk_url
            print(f"Ссылка на ВКонтакте: {vk_url}")
        else:
            profile_data['vk_link'] = ""
            print("Ссылка на ВКонтакте не найдена или некорректна")
    except Exception as e:
        print(f"Ошибка при получении ссылки на ВКонтакте: {e}")
        profile_data['vk_link'] = ""
    
    return profile_data

def extract_vk_name(driver):
    wait = WebDriverWait(driver, 5)  # уменьшаем с 8 до 5 секунд
    try:
        # Ждём появления имени на странице ВК
        name_element = driver.find_element(By.CSS_SELECTOR, "#owner_page_name")
        
        # Получаем текст элемента (ФИ из ВК)
        vk_name = name_element.text.strip()
        
        # Очищаем текст от лишней информации (убираем "заходила три часа назад" и т.д.)
        # Берём только первую часть до первого span с дополнительной информацией
        if 'заходила' in vk_name or 'заходил' in vk_name:
            vk_name = vk_name.split('заходил')[0].strip()
        
        print(f"ФИ из ВКонтакте: {vk_name}")
        return vk_name
        
    except Exception as e:
        print(f"Ошибка при получении ФИ из ВКонтакте: {e}")
        return ""

def compare_and_update_names(platform_name, vk_name):
    """Сравнивает ФИ с платформы и ВК, возвращает ФИ из ВК если они отличаются"""
    if platform_name and vk_name and platform_name != vk_name:
        print(f"ФИ отличаются! Платформа: '{platform_name}', ВК: '{vk_name}'")
        return vk_name
    elif platform_name and vk_name and platform_name == vk_name:
        print(f"ФИ одинаковые: '{platform_name}'")
        return ""
    else:
        print("Не удалось сравнить ФИ")
        return ""

def go_to_vk_and_compare_names(driver, vk_url, platform_name):
    """Открывает страницу ВК в той же вкладке, сравнивает ФИ и возвращается назад на профиль."""
    if not vk_url:
        print("Ссылка на ВК не найдена")
        return ""
    try:
        profile_url = driver.current_url
        driver.get(vk_url)
        # Ждём только имя владельца страницы
        wait = WebDriverWait(driver, 2)  # уменьшаем с 3 до 2 секунд
        vk_name = extract_vk_name(driver)
        different_name = compare_and_update_names(platform_name, vk_name)
        # Возвращаемся на профиль (быстрее, чем управление вкладками)
        try:
            driver.get(profile_url)
        except Exception:
            pass
        return different_name
    except Exception as e:
        print(f"Ошибка при работе с ВК: {e}")
        try:
            driver.get(profile_url)
        except Exception:
            pass
        return ""

def update_excel_with_homework_data(homework_number, cloud_links, student_data, profile_data=None, vk_different_name=""):
    filename = "homework_data.xlsx"
    
    # Убираем детальное логирование для ускорения
    
    # Дополнительная проверка перед сохранением
    mark = student_data.get('mark', '')
    if mark == '' or mark is None:
        print(f"❌ КРИТИЧЕСКАЯ ОШИБКА: Попытка сохранить данные без баллов для ДЗ №{homework_number}")
        return False
    
    if mark >= 2 and not cloud_links:
        print(f"❌ КРИТИЧЕСКАЯ ОШИБКА: Баллов {mark} ≥ 2, но ссылка на облако не найдена для ДЗ №{homework_number}")
        return False
    
    print(f"✅ Данные прошли финальную проверку, сохраняем...")
    
    try:
        # Загружаем существующую таблицу
        df = pd.read_excel(filename)
        # Убираем детальное логирование для ускорения
    except FileNotFoundError:
        # Если файл не найден, создаём новый
        # Убираем детальное логирование для ускорения
        columns = [
            'КО', 'Группа', 'ФИ студента с платформы', 'ВХ Анкета (У1)',
            'Ссылка на ДЗ №1', 'Комментарий к ДЗ №1', 'Оценка по ДЗ №1',
            'Ссылка на ДЗ №2', 'Комментарий к ДЗ №2', 'Оценка по ДЗ №2',
            'Ссылка на профиль на платформе', 'Ссылка на страницу ВКонтакте',
            'ФИ из ВКонтакте (если отличается)', 'Ссылка на сообщество',
            'Количество подписчиков', 'Сумма'
        ]
        df = pd.DataFrame(columns=columns)
    
    # Ищем существующую строку с таким же ФИ и ссылкой на профиль
    existing_row_index = None
    current_full_name = student_data.get('full_name', '')
    current_profile_link = student_data.get('profile_link', '')
    
    # Убираем детальное логирование для ускорения
    
    for index, row in df.iterrows():
        existing_full_name = str(row.get('ФИ студента с платформы', ''))
        existing_profile_link = str(row.get('Ссылка на профиль на платформе', ''))
        
        # Сравниваем ФИ и ссылку на профиль
        if (current_full_name == existing_full_name and 
            current_profile_link == existing_profile_link and 
            current_full_name != '' and current_profile_link != ''):
            existing_row_index = index
            # Убираем детальное логирование для ускорения
            break
    
    if existing_row_index is not None:
        # Обновляем существующую строку
        print(f"Найдена существующая запись для студента: {current_full_name}")
        
        # Заполняем данные в зависимости от номера ДЗ
        # Сначала сохраняем баллы (всегда)
        if homework_number == 1:
            df.at[existing_row_index, 'Оценка по ДЗ №1'] = student_data.get('mark', '')
            print(f"Сохранены баллы за ДЗ №1: {student_data.get('mark', '')}")
        elif homework_number == 2:
            df.at[existing_row_index, 'Оценка по ДЗ №2'] = student_data.get('mark', '')
            print(f"Сохранены баллы за ДЗ №2: {student_data.get('mark', '')}")
        else:
            print(f"❌ Неизвестный номер ДЗ: {homework_number}")
        
        # Затем сохраняем ссылку на облако, если она есть
        if cloud_links:
            cloud_link = process_cloud_links(cloud_links)
            link_col = f'Ссылка на ДЗ №{homework_number}'
            existing_link = df.at[existing_row_index, link_col]
            
            # Проверяем, изменились ли ссылки
            if existing_link != cloud_link:
                if homework_number == 1:
                    df.at[existing_row_index, 'Ссылка на ДЗ №1'] = cloud_link
                elif homework_number == 2:
                    df.at[existing_row_index, 'Ссылка на ДЗ №2'] = cloud_link
                
                if existing_link and existing_link != '':
                    print(f"🔄 Обновлена ссылка на облако (было: {existing_link}, стало: {cloud_link})")
                else:
                    print(f"✅ Добавлена ссылка на облако: {cloud_link}")
            else:
                print(f"ℹ️ Ссылка на облако не изменилась: {cloud_link}")
        else:
            print("Ссылки на облако не найдены, но баллы сохранены")
        
        # Обновляем данные из профиля, если они переданы
        if profile_data:
            if profile_data.get('group'):
                df.at[existing_row_index, 'Группа'] = profile_data['group']
            if profile_data.get('vk_link'):
                df.at[existing_row_index, 'Ссылка на страницу ВКонтакте'] = profile_data['vk_link']
        
        # Обновляем ФИ из ВК, если оно отличается
        if vk_different_name:
            df.at[existing_row_index, 'ФИ из ВКонтакте (если отличается)'] = vk_different_name
        
        # Обновляем ФИ и ссылку на профиль, если они пустые
        if not df.at[existing_row_index, 'ФИ студента с платформы']:
            df.at[existing_row_index, 'ФИ студента с платформы'] = current_full_name
        if not df.at[existing_row_index, 'Ссылка на профиль на платформе']:
            df.at[existing_row_index, 'Ссылка на профиль на платформе'] = current_profile_link
            
    else:
        # Создаём новую строку
        # Убираем детальное логирование для ускорения
        student_number = len(df) + 1
        new_row = pd.Series(index=df.columns)
        
        # Заполняем данные
        new_row['КО'] = student_number
        new_row['ФИ студента с платформы'] = current_full_name
        new_row['Ссылка на профиль на платформе'] = current_profile_link
        
        # Заполняем данные из профиля, если они переданы
        if profile_data:
            new_row['Группа'] = profile_data.get('group', '')
            new_row['Ссылка на страницу ВКонтакте'] = profile_data.get('vk_link', '')
        
        # Заполняем ФИ из ВК, если оно отличается
        if vk_different_name:
            new_row['ФИ из ВКонтакте (если отличается)'] = vk_different_name
        
        # Заполняем баллы в зависимости от номера ДЗ (всегда)
        if homework_number == 1:
            new_row['Оценка по ДЗ №1'] = student_data.get('mark', '')
            print(f"Сохранены баллы за ДЗ №1: {student_data.get('mark', '')}")
        elif homework_number == 2:
            new_row['Оценка по ДЗ №2'] = student_data.get('mark', '')
            print(f"Сохранены баллы за ДЗ №2: {student_data.get('mark', '')}")
        else:
            print(f"❌ Неизвестный номер ДЗ: {homework_number}")
        
        # Заполняем ссылку на ДЗ в зависимости от номера (только если есть)
        if cloud_links:
            cloud_link = process_cloud_links(cloud_links)
            if homework_number == 1:
                new_row['Ссылка на ДЗ №1'] = cloud_link
            elif homework_number == 2:
                new_row['Ссылка на ДЗ №2'] = cloud_link
            print(f"Сохранена ссылка на облако: {cloud_link}")
        else:
            print("Ссылки на облако не найдены, но баллы сохранены")
        
        # Добавляем строку в DataFrame
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        print(f"Добавлен новый студент №{student_number}: {current_full_name}")
    
    # Сохраняем обновлённую таблицу
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Данные')
        
        # Настраиваем ширину колонок
        worksheet = writer.sheets['Данные']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = min(max(max_length + 2, 15), 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f"Данные сохранены в таблицу: {filename}")

def go_to_student_profile(driver, profile_url):
    # Переходим в профиль студента в той же вкладке
    driver.get(profile_url)
    time.sleep(0.5)  # уменьшаем с 0.8 до 0.5 секунды
    print(f"Перешли в профиль студента: {profile_url}")
    
    # Извлекаем данные из профиля
    profile_data = extract_profile_data(driver)
    return profile_data

def get_homework_number_and_fill_data(driver):
    wait = WebDriverWait(driver, 5)  # уменьшаем с 8 до 5 секунд
    try:
        header = driver.find_element(By.CSS_SELECTOR, "#content .homework-chat-header-left h1")
        header_text = header.text.strip()
        print(f"Заголовок ДЗ: {header_text}")
        homework_number = None
        
        # Улучшенная логика определения номера ДЗ
        if "Домашнее задание №1" in header_text or "№1" in header_text:
            homework_number = 1
            print("✅ Определено как ДЗ №1")
        elif "Домашнее задание №2" in header_text or "№2" in header_text or "Оглянитесь по сторонам" in header_text:
            homework_number = 2
            print("✅ Определено как ДЗ №2")
        else:
            print(f"❌ Неизвестный заголовок ДЗ: {header_text}")
            # Попробуем найти номер в тексте заголовка
            import re
            number_match = re.search(r'№(\d+)', header_text)
            if number_match:
                homework_number = int(number_match.group(1))
                print(f"✅ Найден номер ДЗ в заголовке: №{homework_number}")
            else:
                print("❌ Не удалось определить номер ДЗ")
            return None
        
        print(f"Определён номер ДЗ: {homework_number}")

        # Ищем ссылки на облако и сразу удаляем себя из проверяющих
        student_data = extract_student_data(driver)
        print(f"📊 Извлечены данные студента: {student_data.get('full_name', 'N/A')}, баллы: {student_data.get('mark', 'N/A')}")
        
        # Определяем, нужно ли искать ссылки на облако
        mark = student_data.get('mark', '')
        if mark == '' or mark is None:
            print(f"❌ ОШИБКА: Не удалось извлечь баллы студента")
            return homework_number
        
        if mark >= 2:
            print(f"🔍 Ищем ссылки на облако (баллов {mark} ≥ 2)...")
            cloud_links = extract_cloud_links(driver)
        else:
            print(f"🔍 Ссылки на облако не требуются (баллов {mark} < 2)")
            cloud_links = []
        
        print("Удаляем себя из проверяющих после поиска ссылок на облако...")
        try:
            remove_from_reviewers(driver)
            print("✅ Успешно удалён из проверяющих после поиска ссылок")
        except Exception as e:
            print(f"❌ Ошибка при удалении из проверяющих: {e}")
        
        # Валидируем данные студента
        print(f"\n🔍 ВАЛИДАЦИЯ ДАННЫХ ДЛЯ ДЗ №{homework_number}:")
        if not validate_student_data(student_data, cloud_links, homework_number):
            print(f"❌ Валидация не пройдена для ДЗ №{homework_number}. Пропускаем сохранение.")
            return homework_number

        # Проверяем, обработана ли уже домашка для этого студента
        already_processed = check_homework_already_processed(homework_number, student_data, check_links_only=True)
        
        if already_processed:
            print(f"🔍 Домашка №{homework_number} для {student_data.get('full_name', 'N/A')} уже обработана")
            print(f"🔍 Всё равно добавляем в проверяющие для проверки ссылок и баллов")
            print(f"🔍 Но пропускаем переходы в профиль и на ВК для ускорения")
            
            # Для уже обработанных студентов не переходим в профиль и на ВК
            profile_data = None
            vk_different_name = ""
        else:
            print(f"🔍 Домашка №{homework_number} для {student_data.get('full_name', 'N/A')} новая")
            print(f"🔍 Добавляем в проверяющие и переходим в профиль и на ВК для полного сбора данных")
            
            # Переходим в профиль студента и получаем данные
            profile_data = None
            profile_url = student_data.get('profile_link', '')
            if profile_url:
                profile_data = go_to_student_profile(driver, profile_url)

            # Переходим на ВК и сравниваем ФИ — ТОЛЬКО если в этом запуске ещё не проверяли данный профиль
            vk_different_name = ""
            if profile_data and profile_data.get('vk_link'):
                vk_profile_key = profile_data['vk_link'] or profile_url
                if vk_profile_key not in processed_vk_profiles:
                    vk_different_name = go_to_vk_and_compare_names(
                        driver,
                        profile_data['vk_link'],
                        student_data.get('full_name', '')
                    )
                    processed_vk_profiles.add(vk_profile_key)
                else:
                    print("VK уже проверен в этом запуске — пропускаем переход на ВК для ускорения")

        print(f"💾 Сохраняем данные в Excel для ДЗ №{homework_number}...")
        update_excel_with_homework_data(homework_number, cloud_links, student_data, profile_data, vk_different_name)
        
        return homework_number
    except Exception as e:
        print(f"❌ Ошибка при получении заголовка ДЗ: {e}")
        return None

def validate_student_data(student_data, cloud_links, homework_number):
    """Валидирует данные студента согласно бизнес-правилам"""
    errors = []
    warnings = []
    
    # Проверяем обязательность баллов
    mark = student_data.get('mark', '')
    if mark == '' or mark is None:
        errors.append(f"❌ ОШИБКА: Отсутствуют баллы за ДЗ №{homework_number}")
    else:
        print(f"✅ Баллы за ДЗ №{homework_number}: {mark}")
        
        # Проверяем необходимость ссылки на облако
        if mark >= 2:
            if not cloud_links:
                errors.append(f"❌ ОШИБКА: Баллов {mark} ≥ 2, но ссылка на облако не найдена")
            else:
                if len(cloud_links) == 1:
                    print(f"✅ Ссылка на облако найдена для баллов {mark}")
                else:
                    print(f"✅ Найдено {len(cloud_links)} ссылок на облако для баллов {mark}")
        else:
            if cloud_links:
                if len(cloud_links) == 1:
                    warnings.append(f"⚠️ ВНИМАНИЕ: Баллов {mark} < 2, но ссылка на облако найдена")
                else:
                    warnings.append(f"⚠️ ВНИМАНИЕ: Баллов {mark} < 2, но найдено {len(cloud_links)} ссылок на облако")
            else:
                print(f"✅ Баллов {mark} < 2, ссылка на облако не требуется")
    
    # Проверяем наличие ФИ студента
    full_name = student_data.get('full_name', '')
    if not full_name:
        errors.append("❌ ОШИБКА: Отсутствует ФИ студента")
    else:
        print(f"✅ ФИ студента: {full_name}")
    
    # Проверяем наличие ссылки на профиль
    profile_link = student_data.get('profile_link', '')
    if not profile_link:
        errors.append("❌ ОШИБКА: Отсутствует ссылка на профиль студента")
    else:
        print(f"✅ Ссылка на профиль: {profile_link}")
    
    # Выводим все ошибки и предупреждения
    if errors:
        print("\n🚨 ОШИБКИ ВАЛИДАЦИИ:")
        for error in errors:
            print(error)
    
    if warnings:
        print("\n⚠️ ПРЕДУПРЕЖДЕНИЯ:")
        for warning in warnings:
            print(warning)
    
    # Возвращаем True если нет критических ошибок
    return len(errors) == 0

def is_homework_complete(student_data, cloud_links, homework_number):
    """Проверяет, полностью ли обработана домашка согласно бизнес-правилам"""
    mark = student_data.get('mark', '')
    
    # Если нет баллов - домашка не обработана
    if mark == '' or mark is None:
        return False
    
    # Если баллов ≥ 2, то нужна хотя бы одна ссылка на облако
    if mark >= 2:
        return len(cloud_links) > 0
    
    # Если баллов < 2, то ссылка не нужна
    return True

def check_homework_already_processed(homework_number, student_data, check_links_only=False):
    """Проверяет, обработана ли уже данная домашка для данного студента
    
    Args:
        homework_number: номер домашнего задания
        student_data: данные студента
        check_links_only: если True, то для уже обработанных студентов возвращает False,
                         чтобы можно было проверить и обновить ссылки
    """
    filename = "homework_data.xlsx"
    
    # Убираем детальное логирование для ускорения
    
    try:
        df = pd.read_excel(filename)
    except FileNotFoundError:
        # Убираем детальное логирование для ускорения
        return False
    
    current_full_name = student_data.get('full_name', '')
    current_profile_link = student_data.get('profile_link', '')
    
    # Ищем студента в таблице
    for index, row in df.iterrows():
        existing_full_name = str(row.get('ФИ студента с платформы', ''))
        existing_profile_link = str(row.get('Ссылка на профиль на платформе', ''))
        
        if (current_full_name == existing_full_name and 
            current_profile_link == existing_profile_link and 
            current_full_name != '' and current_profile_link != ''):
            
            # Проверяем, есть ли уже баллы по этому ДЗ
            mark_col = f'Оценка по ДЗ №{homework_number}'
            existing_mark = row.get(mark_col)
            # Убираем детальное логирование для ускорения
            
            # Проверяем полноту данных согласно бизнес-правилам
            if mark_col in df.columns and pd.notna(existing_mark) and existing_mark != '':
                # Проверяем, есть ли ссылка на облако (если баллов ≥ 2)
                link_col = f'Ссылка на ДЗ №{homework_number}'
                existing_link = row.get(link_col)
                
                if existing_mark >= 2:
                    if existing_link and existing_link != '':
                        # Проверяем, есть ли хотя бы одна ссылка (может быть несколько через разделитель)
                        # Если check_links_only=True, то возвращаем True для уже обработанных
                        if check_links_only:
                            return True
                        return True
                    else:
                        return False
                else:
                    return True
            else:
                return False
    
    return False

def process_all_homeworks_on_page(driver):
    """Обрабатывает все домашки на текущей странице, начиная с последней"""
    wait = WebDriverWait(driver, 20)  # увеличиваем с 15 до 20 секунд
    
    # Ждём появления таблицы с домашками
    wait.until(EC.presence_of_element_located((
        By.CSS_SELECTOR,
        "#react-homeworks > div > div.sc-gFqAkR.kiXkom > table > tbody > tr"
    )))
    
    # Дополнительная пауза для полной отрисовки таблицы
    time.sleep(5)  # увеличиваем с 3 до 5 секунд
    
    # Проверяем, что таблица действительно загрузилась и содержит данные
    try:
        rows = driver.find_elements(By.CSS_SELECTOR, "#react-homeworks > div > div.sc-gFqAkR.kiXkom > table > tbody > tr")
        total_homeworks = len(rows)
        
        if total_homeworks == 0:
            print("⚠️ Таблица загружена, но строки не найдены. Ждём ещё...")
            time.sleep(5)  # увеличиваем с 3 до 5 секунд
            rows = driver.find_elements(By.CSS_SELECTOR, "#react-homeworks > div > div.sc-gFqAkR.kiXkom > table > tbody > tr")
            total_homeworks = len(rows)
        
        print(f"✅ Таблица загружена! Найдено домашек на странице: {total_homeworks}")
        
        if total_homeworks == 0:
            print("❌ Не удалось загрузить домашние задания. Пропускаем страницу.")
            return
            
    except Exception as e:
        print(f"❌ Ошибка при загрузке таблицы: {e}")
        print("⏳ Продолжаем ждать...")
        time.sleep(12)  # увеличиваем с 8 до 12 секунд
        try:
            rows = driver.find_elements(By.CSS_SELECTOR, "#react-homeworks > div > div.sc-gFqAkR.kiXkom > table > tbody > tr")
            total_homeworks = len(rows)
            print(f"✅ После дополнительного ожидания найдено: {total_homeworks} домашек")
        except Exception as e2:
            print(f"❌ Критическая ошибка загрузки таблицы: {e2}")
            return

    main_window = driver.current_window_handle
    print(f"Handle основной вкладки: {main_window}")

    for hw_number in range(total_homeworks, 0, -1):
        try:
            print(f"\n{'='*50}")
            print(f"Обработка домашки {hw_number} из {total_homeworks}")
            print(f"{'='*50}")

            driver.switch_to.window(main_window)
            time.sleep(0.2)  # увеличиваем с 0.1 до 0.2 секунды

            homework_selector = f"#react-homeworks > div > div.sc-gFqAkR.kiXkom > table > tbody > tr:nth-child({hw_number}) > td.sc-eqUAAy.sc-iGgWBj.jjpiPE.gbwPlL > a"
            homework_link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, homework_selector)))
            homework_url = homework_link.get_attribute("href")
            print(f"URL домашки: {homework_url}")

            driver.execute_script("window.open(arguments[0], '_blank');", homework_url)
            time.sleep(0.5)  # увеличиваем с 0.3 до 0.5 секунды

            initial_handles = set(driver.window_handles)
            homework_window = driver.window_handles[-1]
            driver.switch_to.window(homework_window)

            wait.until(EC.presence_of_element_located((
                By.CSS_SELECTOR,
                "#content > div.homework-chat-header > div.homework-chat-header-left > h1"
            )))
            time.sleep(0.5)  # увеличиваем с 0.3 до 0.5 секунды

            # Быстрая проверка: если уже обработано — скипаем (проверка в начале страницы домашки)
            already_processed = False
            try:
                student_data_check = extract_student_data(driver)
                header = driver.find_element(By.CSS_SELECTOR, "#content > div.homework-chat-header > div.homework-chat-header-left > h1")
                header_text = header.text.strip()
                
                # Используем ту же логику, что и в основной функции
                homework_no = None
                if "Домашнее задание №1" in header_text or "№1" in header_text:
                    homework_no = 1
                elif "Домашнее задание №2" in header_text or "№2" in header_text or "Оглянитесь по сторонам" in header_text:
                    homework_no = 2
                else:
                    # Попробуем найти номер в тексте заголовка
                    import re
                    number_match = re.search(r'№(\d+)', header_text)
                    if number_match:
                        homework_no = int(number_match.group(1))
                
                # Проверяем, обработана ли домашка, но с возможностью обновления ссылок
                if homework_no and check_homework_already_processed(homework_no, student_data_check, check_links_only=True):
                    print(f"Домашка №{homework_no} для {student_data_check.get('full_name','')} уже есть — пропускаем обработку данных.")
                    already_processed = True
                else:
                    pass  # Убираем детальное логирование для ускорения
            except Exception as e:
                # Убираем детальное логирование для ускорения
                pass

            # Назначаем проверяющего и собираем данные (если не обработано)
            if not already_processed:
                try:
                    process_homework_page(driver)
                    print("Успешно добавлены в проверяющие")
                except Exception as e:
                    print(f"Ошибка при добавлении в проверяющие: {e}")

                try:
                    homework_number = get_homework_number_and_fill_data(driver)
                    if homework_number:
                        print(f"Успешно обработана домашка №{homework_number}")
                    else:
                        print("Не удалось определить номер домашки")
                except Exception as e:
                    print(f"Ошибка при извлечении данных: {e}")
            else:
                # Если домашка уже обработана, всё равно добавляем в проверяющие для проверки ссылок
                print("Домашка уже обработана, но всё равно добавляем в проверяющие для проверки ссылок...")
                try:
                    process_homework_page(driver)
                    print("✅ Успешно добавлены в проверяющие (для уже обработанной домашки)")
                except Exception as e:
                    print(f"❌ Ошибка при добавлении в проверяющие: {e}")
                
                # Обрабатываем данные для уже обработанной домашки
                try:
                    homework_number = get_homework_number_and_fill_data(driver)
                    if homework_number:
                        print(f"✅ Успешно обновлены данные для домашки №{homework_number}")
                    else:
                        print("Не удалось определить номер домашки")
                except Exception as e:
                    print(f"Ошибка при обновлении данных: {e}")

            # Закрываем вкладку с домашкой и возвращаемся к списку
            driver.close()
            driver.switch_to.window(main_window)
            time.sleep(0.2)  # увеличиваем с 0.1 до 0.2 секунды

        except Exception as e:
            print(f"Ошибка при обработке домашки №{hw_number}: {e}")
            try:
                for handle in driver.window_handles:
                    if handle != main_window:
                        driver.switch_to.window(handle)
                        driver.close()
                driver.switch_to.window(main_window)
            except Exception as cleanup_error:
                print(f"Ошибка при очистке вкладок: {cleanup_error}")
            continue

    print("\nОбработка всех домашек на странице завершена")

def process_all_pages(driver):
    """Обрабатывает все страницы, начиная с последней"""
    page_number = 1
    
    while True:
        print(f"\n" + "="*60)
        print(f"ОБРАБОТКА СТРАНИЦЫ №{page_number}")
        print(f"="*60)
        
        # Обрабатываем все домашки на текущей странице
        process_all_homeworks_on_page(driver)
        
        # Пытаемся перейти на предыдущую страницу
        try:
            # Ждём загрузки пагинации
            wait = WebDriverWait(driver, 15)  # увеличиваем с 10 до 15 секунд
            wait.until(EC.presence_of_element_located((
                By.CSS_SELECTOR,
                "#react-homeworks > div > div.sc-gFqAkR.kiXkom > div > button.r-button.button-pagination.active"
            )))
            
            # Находим активную страницу
            active_page_btn = driver.find_element(By.CSS_SELECTOR, 
                "#react-homeworks > div > div.sc-gFqAkR.kiXkom > div > button.r-button.button-pagination.active")
            
            # Получаем все кнопки пагинации
            all_page_buttons = driver.find_elements(By.CSS_SELECTOR, 
                "#react-homeworks > div > div.sc-gFqAkR.kiXkom > div > button.r-button.button-pagination")
            
            # Находим индекс активной кнопки
            active_index = -1
            for i, btn in enumerate(all_page_buttons):
                if btn == active_page_btn:
                    active_index = i
                    break
            
            # Проверяем, есть ли предыдущая страница
            if active_index > 0:  # если не первая кнопка
                prev_page_btn = all_page_buttons[active_index - 1]
                
                print(f"\nТекущая страница: {active_page_btn.text}")
                print(f"Переходим на предыдущую страницу: {prev_page_btn.text}")
                
                prev_page_btn.click()
                
                # Ждём загрузки новой страницы
                # Убираем детальное логирование для ускорения
                time.sleep(15)  # увеличиваем с 8 до 15 секунд
                
                # Проверяем, что страница загрузилась
                try:
                    wait.until(EC.presence_of_element_located((
                        By.CSS_SELECTOR,
                        "#react-homeworks > div > div.sc-gFqAkR.kiXkom > table > tbody > tr"
                    )))
                    print("✅ Новая страница загружена")
                except Exception as load_error:
                    print(f"⚠️ Страница загружается медленно: {load_error}")
                    time.sleep(8)  # увеличиваем с 3 до 8 секунд
                
                page_number += 1
                # Убираем детальное логирование для ускорения
            else:
                print("\nДостигнута первая страница. Больше страниц нет.")
                break
                
        except Exception as e:
            print(f"\nОшибка при переходе на предыдущую страницу: {e}")
            print("Возможно, это была последняя страница.")
            break
    
    print(f"\n" + "="*60)
    print(f"ОБРАБОТКА ЗАВЕРШЕНА! Обработано страниц: {page_number}")
    print(f"="*60)

def go_to_last_page(driver):
    """Переходит на последнюю страницу со списком домашек"""
    wait = WebDriverWait(driver, 25)  # увеличиваем с 20 до 25 секунд
    # Ждём появления кнопки последней страницы
    last_page_btn = wait.until(EC.element_to_be_clickable((
        By.CSS_SELECTOR,
        "#react-homeworks > div > div.sc-gFqAkR.kiXkom > div > button.r-button.button-pagination.boundary > span"
    )))
    
    last_page_btn.click()
    
    # Ждём загрузки таблицы с домашками
    wait.until(EC.presence_of_element_located((
        By.CSS_SELECTOR,
        "#react-homeworks > div > div.sc-gFqAkR.kiXkom > table > tbody > tr"
    )))
    
    # Увеличиваем паузу для полной отрисовки
    time.sleep(20)  # увеличиваем с 15 до 20 секунд
    
    # Проверяем, что таблица действительно загрузилась
    try:
        rows = driver.find_elements(By.CSS_SELECTOR, "#react-homeworks > div > div.sc-gFqAkR.kiXkom > table > tbody > tr")
        print(f"✅ Таблица загружена! Найдено {len(rows)} домашних заданий на последней странице")
        
        # Дополнительная проверка - убеждаемся, что строки содержат данные
        if rows:
            first_row = rows[0]
            try:
                # Проверяем, что в первой строке есть ссылка на домашку
                homework_link = first_row.find_element(By.CSS_SELECTOR, "td.sc-eqUAAy.sc-iGgWBj.jjpiPE.gbwPlL > a")
                if homework_link:
                    print("✅ Первая строка содержит ссылку на домашку - таблица готова")
                else:
                    time.sleep(12)  # увеличиваем с 8 до 12 секунд
            except Exception as link_error:
                time.sleep(15)  # увеличиваем с 10 до 15 секунд
        else:
            time.sleep(15)  # увеличиваем с 10 до 15 секунд
            rows = driver.find_elements(By.CSS_SELECTOR, "#react-homeworks > div > div.sc-gFqAkR.kiXkom > table > tbody > tr")
            print(f"✅ После дополнительного ожидания найдено: {len(rows)} строк")
            
    except Exception as e:
        time.sleep(20)  # увеличиваем с 15 до 20 секунд
        
        # Финальная попытка
        try:
            rows = driver.find_elements(By.CSS_SELECTOR, "#react-homeworks > div > div.sc-gFqAkR.kiXkom > table > tbody > tr")
            print(f"✅ Финальная проверка: найдено {len(rows)} домашних заданий")
        except Exception as e2:
            pass
    
    print("✅ Перешли на последнюю страницу и готова к обработке")

def process_cloud_links(cloud_links):
    """Обрабатывает множественные ссылки на облако, возвращает строку для сохранения"""
    if not cloud_links:
        return ""
    
    if len(cloud_links) == 1:
        return cloud_links[0]
    
    # Если ссылок несколько, объединяем их через разделитель
    print(f"🔍 Найдено {len(cloud_links)} ссылок на облако:")
    for i, link in enumerate(cloud_links, 1):
        print(f"   {i}. {link}")
    
    # Объединяем все ссылки через " | " (вертикальная черта)
    combined_links = " | ".join(cloud_links)
    print(f"✅ Все ссылки объединены в одну строку")
    
    return combined_links

def is_valid_cloud_link(href):
    """Проверяет, является ли ссылка валидной ссылкой на облако"""
    if not href or not href.startswith('http'):
        return False
    
    # Исключаем только явно запрещённые домены
    forbidden_domains = ["vk.cc", "vk.me", "authors.vk.company"]
    if any(domain in href for domain in forbidden_domains):
        return False
    
    # Добавляем ВСЕ ссылки на vk.com (включая те, где внутри облачные сервисы)
    if "vk.com" in href:
        return True
    
    # Исключаем обычные mail.ru (но оставляем cloud.mail.ru)
    if "mail.ru" in href and "cloud.mail.ru" not in href:
        return False
    
    # Добавляем ВСЕ остальные HTTP ссылки
    return True

if __name__ == "__main__":
    create_excel_table()
    driver = open_vk_homework_page()
    go_to_last_page(driver)  # Переходим на последнюю страницу
    process_all_pages(driver)  # Обрабатываем все страницы
    print("Скрипт работает. Закройте браузер для завершения.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass
