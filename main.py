"""
课程表生成器 - 主程序入口
==========================
本程序用于从Excel文件中提取课程信息，并生成教师个人课程表的PDF文件。

功能特点：
1. 交互式命令行界面，支持分页浏览教师列表
2. 支持按姓名搜索教师
3. 自动生成格式化的PDF课程表

使用方法：
    python main.py
"""

# 导入课程表处理器类
from schedule_processor import ScheduleProcessor
# 导入操作系统模块，用于文件路径操作
import os


def print_header():
    """
    打印程序标题头
    
    在程序启动时显示欢迎信息，包含俄语和英语标题
    """
    print("=" * 60)
    print(" " * 15 + "Генератор расписания")  # 俄语：课程表生成器
    print(" " * 10 + "Schedule Generator")    # 英语：课程表生成器
    print("=" * 60)


def print_teachers(teachers, page=1, per_page=10):
    """
    分页打印教师列表
    
    参数:
        teachers (list): 教师姓名列表
        page (int): 当前页码，默认为第1页
        per_page (int): 每页显示数量，默认为10条
    """
    # 计算当前页的起始索引
    start_idx = (page - 1) * per_page
    # 计算当前页的结束索引（不超过列表总长度）
    end_idx = min(start_idx + per_page, len(teachers))
    
    # 打印页码信息
    print(f"\nСписок преподавателей (страница {page} из {(len(teachers) + per_page - 1) // per_page}):")
    print("-" * 60)
    # 遍历并打印当前页的教师
    for i in range(start_idx, end_idx):
        print(f"{i+1:3d}. {teachers[i]}")
    print("-" * 60)


def main():
    """
    主函数 - 程序入口点
    
    处理用户交互流程：
    1. 加载Excel课程表文件
    2. 解析课程数据
    3. 提供交互式菜单供用户选择教师
    4. 生成选中教师的PDF课程表
    """
    # 打印程序标题
    print_header()
    
    # Excel文件路径
    excel_file = 'ITsTiM_Raspisanie_2_polugodie_25-26_bak__pechat.xlsx'
    
    # 检查文件是否存在
    if not os.path.exists(excel_file):
        print(f"\nОшибка: файл не найден {excel_file}")  # 错误：文件未找到
        return
    
    # 加载并解析课程表
    print("\nЗагрузка расписания...")  # 正在加载课程表...
    processor = ScheduleProcessor(excel_file)  # 创建处理器实例
    processor.load_data()      # 加载Excel数据
    processor.parse_schedule() # 解析课程安排
    
    # 获取所有教师列表
    teachers = processor.get_teachers()
    print(f"Успешно загружено! Найдено {len(teachers)} преподавателей")  # 成功加载！找到X位教师
    
    # 分页控制变量
    current_page = 1   # 当前页码
    per_page = 10      # 每页显示数量
    
    # 主循环 - 处理用户输入
    while True:
        # 显示当前页的教师列表
        print_teachers(teachers, current_page, per_page)
        
        # 显示操作选项菜单
        print("\nОпции:")  # 选项
        print("  Введите номер - выбрать преподавателя")  # 输入编号 - 选择教师
        print("  n/N     - следующая страница")           # 下一页
        print("  p/P     - предыдущая страница")          # 上一页
        print("  s/S     - поиск преподавателя")          # 搜索教师
        print("  q/Q     - выход")                        # 退出
        
        # 获取用户输入
        choice = input("\nВыберите > ").strip()  # 请选择 >
        
        # 处理退出命令
        if choice.lower() == 'q':
            print("\nСпасибо за использование!")  # 感谢使用！
            break
        
        # 处理下一页命令
        elif choice.lower() == 'n':
            if current_page < (len(teachers) + per_page - 1) // per_page:
                current_page += 1
            else:
                print("\nЭто последняя страница")  # 这是最后一页
        
        # 处理上一页命令
        elif choice.lower() == 'p':
            if current_page > 1:
                current_page -= 1
            else:
                print("\nЭто первая страница")  # 这是第一页
        
        # 处理搜索命令
        elif choice.lower() == 's':
            # 获取搜索关键词
            search_term = input("Введите имя преподавателя (поддерживается частичное совпадение) > ").strip()
            # 执行模糊搜索（支持部分匹配）
            matched = [t for t in teachers if search_term.lower() in t.lower()]
            
            if matched:
                # 显示搜索结果
                print(f"\nНайдено {len(matched)} преподавателей:")  # 找到X位教师
                for i, teacher in enumerate(matched, 1):
                    print(f"{i}. {teacher}")
                
                # 让用户从搜索结果中选择
                sub_choice = input("\nВыберите номер (или Enter для возврата) > ").strip()
                if sub_choice.isdigit() and 1 <= int(sub_choice) <= len(matched):
                    selected_teacher = matched[int(sub_choice) - 1]
                    generate_schedule(processor, selected_teacher)
            else:
                print("\nПреподаватели не найдены")  # 未找到教师
        
        # 处理数字选择
        elif choice.isdigit():
            idx = int(choice) - 1  # 转换为0-based索引
            if 0 <= idx < len(teachers):
                selected_teacher = teachers[idx]
                generate_schedule(processor, selected_teacher)
            else:
                print("\nНеверный номер")  # 无效编号
        
        # 处理无效输入
        else:
            print("\nНеверная опция")  # 无效选项


def generate_schedule(processor, teacher_name):
    """
    为指定教师生成PDF课程表
    
    参数:
        processor (ScheduleProcessor): 课程表处理器实例
        teacher_name (str): 教师姓名
    """
    print(f"\nВыбрано: {teacher_name}")  # 已选择
    
    # 获取该教师的课程安排
    schedule = processor.get_teacher_schedule(teacher_name)
    print(f"Всего занятий: {len(schedule)}")  # 总课程数
    
    # 确认是否生成PDF
    confirm = input("\nПодтвердить генерацию PDF? (y/n) > ").strip().lower()
    if confirm == 'y':
        # 生成输出文件名（替换空格和点号）
        output_file = f"schedule_{teacher_name.replace(' ', '_').replace('.', '_')}.pdf"
        
        # 调用处理器生成PDF
        if processor.export_to_pdf(teacher_name, output_file):
            print(f"\n✓ PDF-файл успешно создан: {output_file}")  # PDF文件创建成功
            print(f"  Расположение: {os.path.abspath(output_file)}")  # 文件位置
        else:
            print("\n✗ Ошибка генерации PDF")  # PDF生成错误
    else:
        print("\nОтменено")  # 已取消


# 程序入口点
if __name__ == "__main__":
    main()
