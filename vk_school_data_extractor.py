import pandas as pd
import re
from urllib.parse import urlparse

def extract_profile_id_from_url(profile_url):
    """Извлекает ID профиля из URL authors.vk.company"""
    if not profile_url or pd.isna(profile_url):
        return None
    
    # Примеры:
    # https://authors.vk.company/profile/a.khaliullina-indradzh/
    # https://authors.vk.company/cabinet/k.adamova/
    # Извлекаем: a.khaliullina-indradzh или k.adamova
    match = re.search(r'/(?:profile|cabinet)/([^/]+)/?$', str(profile_url))
    if match:
        return match.group(1)
    return None



def find_student_in_vk_school(homework_data_row, vk_school_df):
    """Ищет студента в таблице Школы авторов VK"""
    
    # Получаем данные из homework_data
    student_name = str(homework_data_row['ФИ студента с платформы']).strip()
    profile_url = str(homework_data_row['Ссылка на профиль на платформе']).strip()
    
    if pd.isna(student_name) or student_name == '' or pd.isna(profile_url) or profile_url == '':
        print(f"❌ Пропускаем строку - отсутствуют ФИ или ссылка на профиль")
        return None
    
    print(f"🔍 Ищем студента: {student_name}")
    print(f"🔍 Профиль: {profile_url}")
    
    # Извлекаем ID профиля (убираем /profile/ или /cabinet/)
    profile_id = extract_profile_id_from_url(profile_url)
    if not profile_id:
        print(f"❌ Не удалось извлечь ID профиля из URL: {profile_url}")
        return None
    
    print(f"🔍 ID профиля: {profile_id}")
    
    # Ищем в таблице Школы авторов VK только по ID профиля
    for index, vk_row in vk_school_df.iterrows():
        vk_name = str(vk_row.iloc[2]).strip()  # Столбец C (ФИ)
        vk_profile = str(vk_row.iloc[3]).strip()  # Столбец D (ссылка на профиль)
        
        # Проверяем совпадение только по ID профиля
        profile_matches = profile_id in vk_profile
        
        if profile_matches:
            print(f"✅ Найден студент в строке {index + 1}")
            print(f"✅ ФИ в Школе авторов VK: {vk_name}")
            print(f"✅ Ссылка в Школе авторов VK: {vk_profile}")
            
            # Извлекаем нужные данные
            u1_data = vk_row.iloc[6]  # Столбец G (У1 - ВХ анкета)
            u7_38_data = vk_row.iloc[9]  # Столбец J (У7/38)
            u7_5_data = vk_row.iloc[10]  # Столбец K (У7/5)
            
            print(f"✅ У1 (ВХ анкета): {u1_data}")
            print(f"✅ У7/38: {u7_38_data}")
            print(f"✅ У7/5: {u7_5_data}")
            
            return {
                'u1': u1_data,
                'u7_38': u7_38_data,
                'u7_5': u7_5_data,
                'vk_school_row': index + 1,
                'vk_name': vk_name
            }
    
    print(f"❌ Студент {student_name} не найден в Школе авторов VK")
    return None

def update_homework_data():
    """Основная функция обновления данных"""
    
    print("🚀 Начинаем обновление данных из Школы авторов VK...")
    
    try:
        # Загружаем таблицу homework_data
        print("📖 Загружаем homework_data.xlsx...")
        homework_df = pd.read_excel('homework_data.xlsx')
        print(f"✅ Загружено {len(homework_df)} строк из homework_data.xlsx")
        
        # Загружаем таблицу Школы авторов VK
        print("📖 Загружаем Школа авторов VK ТБ (ТБ) 2025-08-17.xlsx...")
        vk_school_df = pd.read_excel('Школа авторов VK ТБ (ТБ) 2025-08-17.xlsx')
        print(f"✅ Загружено {len(vk_school_df)} строк из Школы авторов VK")
        
        # Показываем структуру таблиц
        print(f"\n📊 Структура homework_data.xlsx:")
        print(f"   Столбец C: {homework_df.columns[2]} (ФИ студента)")
        print(f"   Столбец D: {homework_df.columns[3]} (ВХ Анкета У1)")
        print(f"   Столбец K: {homework_df.columns[10]} (У7/5)")
        print(f"   Столбец L: {homework_df.columns[11]} (У7/38)")
        print(f"   Столбец M: {homework_df.columns[12]} (Ссылка на профиль)")
        
        print(f"\n📊 Структура Школы авторов VK:")
        print(f"   Столбец C: {vk_school_df.columns[2]} (ФИ)")
        print(f"   Столбец D: {vk_school_df.columns[3]} (Профиль)")
        print(f"   Столбец G: {vk_school_df.columns[6]} (У1 - ВХ анкета)")
        print(f"   Столбец J: {vk_school_df.columns[9]} (У7/38)")
        print(f"   Столбец K: {vk_school_df.columns[10]} (У7/5)")
        
        print(f"\n🔍 Форматы ссылок:")
        print(f"   homework_data: /profile/ID/")
        print(f"   Школа авторов VK: /cabinet/ID/")
        print(f"   Поиск по ID профиля (без /profile/ и /cabinet/)")
        
        print(f"\n🔍 Поиск:")
        print(f"   Поиск только по ID профиля (без /profile/ и /cabinet/)")
        print(f"   ФИ не учитывается при поиске")
        
        # Счётчики
        found_count = 0
        not_found_count = 0
        updated_count = 0
        
        # Список студентов, которые не были найдены
        not_found_students = []
        
        # Обрабатываем каждую строку в homework_data
        for index, row in homework_df.iterrows():
            print(f"\n{'='*60}")
            print(f"Обработка строки {index + 1} из {len(homework_df)}")
            print(f"{'='*60}")
            
            # Ищем студента в Школе авторов VK
            student_data = find_student_in_vk_school(row, vk_school_df)
            
            if student_data:
                found_count += 1
                
                # Обновляем данные в homework_data
                try:
                    # Столбец C (ФИ студента) - заменяем на ФИ из Школы авторов VK
                    homework_df.iloc[index, 2] = student_data['vk_name']
                    
                    # Столбец D (ВХ Анкета У1)
                    homework_df.iloc[index, 3] = student_data['u1']
                    
                    # Столбец K (У7/5)
                    homework_df.iloc[index, 10] = student_data['u7_5']
                    
                    # Столбец L (У7/38)
                    homework_df.iloc[index, 11] = student_data['u7_38']
                    
                    updated_count += 1
                    print(f"✅ Данные обновлены в строке {index + 1}")
                    print(f"✅ ФИ обновлен: {student_data['vk_name']}")
                    
                except Exception as e:
                    print(f"❌ Ошибка при обновлении данных: {e}")
            else:
                not_found_count += 1
                # Добавляем в список ненайденных студентов
                student_name = str(row['ФИ студента с платформы']).strip()
                profile_url = str(row['Ссылка на профиль на платформе']).strip()
                not_found_students.append({
                    'name': student_name,
                    'profile': profile_url,
                    'row': index + 1
                })
        
        # Сохраняем обновлённую таблицу
        print(f"\n💾 Сохраняем обновлённую таблицу...")
        with pd.ExcelWriter('homework_data.xlsx', engine='openpyxl') as writer:
            homework_df.to_excel(writer, index=False, sheet_name='Данные')
            
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
        
        print(f"✅ Таблица сохранена!")
        
        # Итоговая статистика
        print(f"\n📊 ИТОГОВАЯ СТАТИСТИКА:")
        print(f"   Всего строк обработано: {len(homework_df)}")
        print(f"   Студентов найдено: {found_count}")
        print(f"   Студентов не найдено: {not_found_count}")
        print(f"   Строк обновлено: {updated_count}")
        
        # Выводим список ненайденных студентов
        if not_found_students:
            print(f"\n❌ СТУДЕНТЫ, КОТОРЫЕ НЕ БЫЛИ НАЙДЕНЫ:")
            print(f"{'='*80}")
            for student in not_found_students:
                print(f"   Строка {student['row']}: {student['name']}")
                print(f"   Профиль: {student['profile']}")
                print(f"   {'-'*60}")
        else:
            print(f"\n✅ Все студенты были успешно найдены!")
        
    except Exception as e:
        print(f"❌ Критическая ошибка: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    update_homework_data()
    print("\n�� Скрипт завершён!") 