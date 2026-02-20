from schedule_processor import ScheduleProcessor
import os

def print_header():
    print("=" * 60)
    print(" " * 15 + "Генератор расписания")
    print(" " * 10 + "Schedule Generator")
    print("=" * 60)

def print_teachers(teachers, page=1, per_page=10):
    start_idx = (page - 1) * per_page
    end_idx = min(start_idx + per_page, len(teachers))
    
    print(f"\nСписок преподавателей (страница {page} из {(len(teachers) + per_page - 1) // per_page}):")
    print("-" * 60)
    for i in range(start_idx, end_idx):
        print(f"{i+1:3d}. {teachers[i]}")
    print("-" * 60)

def main():
    print_header()
    
    excel_file = '22ITsTiM_Raspisanie_2_polugodie_25-26_mag__pechat2.xlsx'
    
    if not os.path.exists(excel_file):
        print(f"\nОшибка: файл не найден {excel_file}")
        return
    
    print("\nЗагрузка расписания...")
    processor = ScheduleProcessor(excel_file)
    processor.load_data()
    processor.parse_schedule()
    
    teachers = processor.get_teachers()
    print(f"Успешно загружено! Найдено {len(teachers)} преподавателей")
    
    current_page = 1
    per_page = 10
    
    while True:
        print_teachers(teachers, current_page, per_page)
        
        print("\nОпции:")
        print("  Введите номер - выбрать преподавателя")
        print("  n/N     - следующая страница")
        print("  p/P     - предыдущая страница")
        print("  s/S     - поиск преподавателя")
        print("  q/Q     - выход")
        
        choice = input("\nВыберите > ").strip()
        
        if choice.lower() == 'q':
            print("\nСпасибо за использование!")
            break
        elif choice.lower() == 'n':
            if current_page < (len(teachers) + per_page - 1) // per_page:
                current_page += 1
            else:
                print("\nЭто последняя страница")
        elif choice.lower() == 'p':
            if current_page > 1:
                current_page -= 1
            else:
                print("\nЭто первая страница")
        elif choice.lower() == 's':
            search_term = input("Введите имя преподавателя (поддерживается частичное совпадение) > ").strip()
            matched = [t for t in teachers if search_term.lower() in t.lower()]
            if matched:
                print(f"\nНайдено {len(matched)} преподавателей:")
                for i, teacher in enumerate(matched, 1):
                    print(f"{i}. {teacher}")
                
                sub_choice = input("\nВыберите номер (или Enter для возврата) > ").strip()
                if sub_choice.isdigit() and 1 <= int(sub_choice) <= len(matched):
                    selected_teacher = matched[int(sub_choice) - 1]
                    generate_schedule(processor, selected_teacher)
            else:
                print("\nПреподаватели не найдены")
        elif choice.isdigit():
            idx = int(choice) - 1
            if 0 <= idx < len(teachers):
                selected_teacher = teachers[idx]
                generate_schedule(processor, selected_teacher)
            else:
                print("\nНеверный номер")
        else:
            print("\nНеверная опция")

def generate_schedule(processor, teacher_name):
    print(f"\nВыбрано: {teacher_name}")
    
    schedule = processor.get_teacher_schedule(teacher_name)
    print(f"Всего занятий: {len(schedule)}")
    
    confirm = input("\nПодтвердить генерацию PDF? (y/n) > ").strip().lower()
    if confirm == 'y':
        output_file = f"schedule_{teacher_name.replace(' ', '_').replace('.', '_')}.pdf"
        if processor.export_to_pdf(teacher_name, output_file):
            print(f"\n✓ PDF-файл успешно создан: {output_file}")
            print(f"  Расположение: {os.path.abspath(output_file)}")
        else:
            print("\n✗ Ошибка генерации PDF")
    else:
        print("\nОтменено")

if __name__ == "__main__":
    main()
