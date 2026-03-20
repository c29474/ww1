"""
课程表处理器模块
================
本模块提供从Excel文件中提取课程信息并生成PDF课程表的核心功能。

主要类：
    ScheduleProcessor - 课程表处理器，负责数据加载、解析和PDF导出

数据流程：
    1. load_data() - 加载Excel文件
    2. parse_schedule() - 解析课程安排
    3. export_to_pdf() - 导出PDF文件

Excel文件格式说明：
    - 第1-2行：标题行
    - 第3行：组名（如 ПМ-11, ИиВТ-21）
    - 第4行：表头（День, Время, ...）
    - 第5行起：课程数据
    - 第1列：星期（如 П О Н Е Д Е Л Ь Н И К）
    - 第2列：时间（如 1-2 с 8.30）
    - 第3列起：各组的课程信息

课程单元格格式：
    课程名称 (类型)
    教师姓名
    教室编号
    
    例如：
    Алгебра и геометрия (лк)
    Игонина Е.В. 4-15
"""

# 数据处理库
import pandas as pd
# 正则表达式模块，用于文本匹配和提取
import re

# PDF生成相关库 - ReportLab
from reportlab.lib.pagesizes import A4           # A4纸张尺寸
from reportlab.lib import colors                  # 颜色定义
from reportlab.platypus import SimpleDocTemplate  # 简单文档模板
from reportlab.platypus import Table              # 表格
from reportlab.platypus import TableStyle         # 表格样式
from reportlab.platypus import Paragraph          # 段落（支持自动换行）
from reportlab.platypus import Spacer             # 间距元素
from reportlab.lib.styles import getSampleStyleSheet    # 获取示例样式
from reportlab.lib.styles import ParagraphStyle          # 段落样式
from reportlab.lib.units import cm                # 单位转换（厘米）
from reportlab.pdfbase import pdfmetrics          # PDF字体度量
from reportlab.pdfbase.ttfonts import TTFont      # TrueType字体支持

# 操作系统模块
import os


class ScheduleProcessor:
    """
    课程表处理器类
    
    负责从Excel文件中提取课程信息，解析课程安排，并生成PDF格式的教师课程表。
    
    属性:
        excel_file (str): Excel文件路径
        df (DataFrame): 加载的Excel数据
        teachers (set): 所有教师姓名集合
        schedule_data (list): 解析后的课程数据列表
        group_grade_map (dict): 列索引到组信息的映射
        
    使用示例:
        processor = ScheduleProcessor('schedule.xlsx')
        processor.load_data()
        processor.parse_schedule()
        processor.export_to_pdf('Иванов И.И.', 'output.pdf')
    """
    
    def __init__(self, excel_file):
        """
        初始化课程表处理器
        
        参数:
            excel_file (str): Excel课程表文件路径
        """
        self.excel_file = excel_file      # Excel文件路径
        self.df = None                    # pandas DataFrame，存储原始数据
        self.teachers = set()             # 教师姓名集合（自动去重）
        self.schedule_data = []           # 解析后的课程数据列表
        self.group_grade_map = {}         # 列索引 -> {group: 组名, grade: 年级}
        
        # 注册PDF字体
        self._register_fonts()
        
    def _register_fonts(self):
        """
        注册PDF字体（私有方法）
        
        注册Arial字体用于PDF生成，支持俄语字符显示。
        在Windows系统中查找系统字体目录下的arial.ttf文件。
        """
        try:
            # 获取Windows字体目录路径
            font_path = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts', 'arial.ttf')
            
            if os.path.exists(font_path):
                # 注册常规字体
                pdfmetrics.registerFont(TTFont('Arial', font_path))
                # 注册粗体字体
                pdfmetrics.registerFont(TTFont('Arial-Bold', font_path.replace('arial.ttf', 'arialbd.ttf')))
            else:
                # 备用路径
                font_path = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts', 'arial.ttf')
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont('Arial', font_path))
        except Exception as e:
            print(f"字体注册警告: {e}")
        
    def load_data(self):
        """
        加载Excel数据
        
        从Excel文件中读取课程表数据，自动选择正确的工作表。
        工作表选择优先级：
        1. Бак_2024-2025（本科生2024-2025学年）
        2. 表6
        3. 第一个工作表
        """
        # 打开Excel文件
        xl = pd.ExcelFile(self.excel_file)
        sheet_names = xl.sheet_names
        
        # 按优先级选择工作表
        if 'Бак_2024-2025' in sheet_names:
            sheet_name = 'Бак_2024-2025'
        elif '表6' in sheet_names:
            sheet_name = '表6'
        else:
            sheet_name = sheet_names[0]
            print(f"Предупреждение: используется лист {sheet_name}")  # 警告：使用工作表
        
        # 读取Excel数据（无表头，保留原始行结构）
        self.df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
        print(f"Успешно загружено расписание, {len(self.df)} строк")  # 成功加载课程表，X行
        
        # 提取组名和年级信息
        self._extract_grade_info()
    
    def _extract_grade_info(self):
        """
        提取组名和年级信息（私有方法）
        
        从Excel第3行（索引2）提取各组名称和对应的年级。
        组名格式示例：ПМ-11（应用数学1年级1组）
        
        数据存储在 group_grade_map 字典中：
        {
            列索引: {
                'group': 'ПМ-11',    # 完整组名
                'grade': 'ПМ-11'     # 年级标识（俄语字母+数字）
            }
        }
        """
        self.group_grade_map = {}
        
        if len(self.df) > 2:
            # 遍历第3行（索引2）的所有列
            for col_idx in range(2, len(self.df.columns)):
                group_name = self.df.iloc[2, col_idx]
                
                # 检查是否为有效的组名
                if pd.notna(group_name) and isinstance(group_name, str):
                    # 从组名中提取年级标识
                    grade = self._extract_grade_from_group(group_name)
                    self.group_grade_map[col_idx] = {
                        'group': group_name,
                        'grade': grade
                    }
        
        print(f"Извлечена информация о {len(self.group_grade_map)} группах")  # 提取了X个组的信息
    
    def _extract_grade_from_group(self, group_name):
        """
        从组名中提取年级标识（私有方法）
        
        俄语组名格式：字母缩写-年级组号
        例如：ПМ-11 -> ПМ-11（应用数学1年级1组）
              ИиВТ-21 -> ИиВТ-21（信息与计算技术2年级1组）
        
        参数:
            group_name (str): 组名
            
        返回:
            str: 年级标识（俄语字母+数字组合）
        """
        # 匹配俄语字母开头，后跟可选连字符/空格和数字的模式
        match = re.search(r'([А-ЯЁа-яё]+[-\s]*\d+)', group_name)
        if match:
            # 移除空格，返回标准化格式
            return match.group(1).replace(' ', '')
        return group_name.strip()
        
    def parse_schedule(self):
        """
        解析课程安排
        
        遍历Excel数据，识别星期和时间，解析每个课程单元格。
        
        解析流程：
        1. 识别星期行（如 "П О Н Е Д Е Л Ь Н И К"）
        2. 识别时间列（如 "1-2 с 8.30"）
        3. 解析课程单元格内容
        
        课程数据存储在 schedule_data 列表中，每条记录包含：
        - day: 星期
        - time: 时间
        - lesson: 课程名称
        - type: 课程类型（лк-讲座, пз-实践课, лб-实验课）
        - teacher: 教师姓名
        - room: 教室编号
        - grade: 年级/组标识
        - group: 完整组名
        """
        self.schedule_data = []
        current_day = None   # 当前星期
        current_time = None  # 当前时间段
        
        # 遍历每一行
        for idx, row in self.df.iterrows():
            day = row[0]   # 第1列：星期
            time = row[1]  # 第2列：时间
            
            # 识别星期（俄语星期名，字母间有空格）
            if pd.notna(day) and 'П О Н Е Д Е Л Ь Н И К' in str(day):
                current_day = 'Понедельник'  # 星期一
            elif pd.notna(day) and 'В Т О Р Н И К' in str(day):
                current_day = 'Вторник'       # 星期二
            elif pd.notna(day) and 'С Р Е Д А' in str(day):
                current_day = 'Среда'         # 星期三
            elif pd.notna(day) and 'Ч Е Т В Е Р Г' in str(day):
                current_day = 'Четверг'       # 星期四
            elif pd.notna(day) and 'П Я Т Н И Ц А' in str(day):
                current_day = 'Пятница'       # 星期五
            elif pd.notna(day) and 'С У Б Б О Т А' in str(day):
                current_day = 'Суббота'       # 星期六
                
            # 更新当前时间
            if pd.notna(time):
                current_time = str(time).replace('\n', ' ')
            
            # 如果已识别星期和时间，开始解析课程单元格
            if current_day and current_time and idx > 0:
                # 遍历第3列起的各组课程
                for col_idx in range(2, len(row)):
                    cell_value = row[col_idx]
                    
                    # 检查单元格是否有有效内容（长度>5字符）
                    if pd.notna(cell_value) and isinstance(cell_value, str) and len(cell_value.strip()) > 5:
                        self._parse_cell(cell_value, current_day, current_time, col_idx)
        
        print(f"Парсинг завершен, найдено {len(self.schedule_data)} записей")  # 解析完成，找到X条记录
        print(f"Всего {len(self.teachers)} преподавателей")  # 共X位教师
        
    def _parse_cell(self, cell_value, day, time, col_idx):
        """
        解析单个课程单元格（私有方法）
        
        课程单元格格式：
            课程名称 (类型)
            教师姓名
            教室编号
            
        例如：
            Алгебра и геометрия (лк)
            Игонина Е.В. 4-15
        
        参数:
            cell_value: 单元格内容
            day: 星期
            time: 时间
            col_idx: 列索引（用于获取组信息）
        """
        cell_value = str(cell_value).strip()
        
        # 提取教师姓名
        # 正则模式：俄语大写字母开头+小写字母+空格+大写字母缩写（1-2个）
        # 例如：Игонина Е.В. 或 Васильева И.И.
        teacher_match = re.search(r'([А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]?\.?)', cell_value)
        
        if teacher_match:
            # 规范化教师姓名（移除多余空格）
            teacher = ' '.join(teacher_match.group(1).split())
            self.teachers.add(teacher)
            
            # 提取教室编号
            # 格式：数字-数字[字母]，如 4-15, 15-305а
            room_match = re.search(r'(\d+-\d+[а-я]?)', cell_value)
            room = room_match.group(1) if room_match else ''
            
            # 提取课程类型
            # лк = лекция（讲座）
            # пз = практическое занятие（实践课）
            # лб = лабораторная работа（实验课）
            lesson_type_match = re.search(r'\((лк|пз|лб)\)', cell_value)
            lesson_type = lesson_type_match.group(1) if lesson_type_match else ''
            
            # 提取课程名称（移除教师姓名和教室编号）
            # 方法：用换行符替换教师姓名，取第一部分
            lesson_name = re.sub(r'\s*[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]?\.?\s*', '\n', cell_value)
            lesson_name = lesson_name.split('\n')[0].strip()
            # 移除末尾的教室编号
            lesson_name = re.sub(r'\s*\d+-\d+[а-я]?\s*$', '', lesson_name)
            # 规范化空格
            lesson_name = re.sub(r'\s+', ' ', lesson_name).strip()
            
            # 获取组信息
            grade_info = self.group_grade_map.get(col_idx, {})
            grade = grade_info.get('grade', '')
            group_name = grade_info.get('group', '')
            
            if not grade:
                grade = ''
            
            # 添加到课程数据列表
            self.schedule_data.append({
                'day': day,           # 星期
                'time': time,         # 时间
                'lesson': lesson_name,  # 课程名称
                'type': lesson_type,    # 课程类型
                'teacher': teacher,     # 教师姓名
                'room': room,           # 教室
                'group_col': col_idx,   # 列索引
                'grade': str(grade) if grade else '',  # 年级/组标识
                'group': group_name    # 完整组名
            })
    
    def get_teachers(self):
        """
        获取所有教师列表
        
        返回:
            list: 按字母排序的教师姓名列表
        """
        return sorted(list(self.teachers))
    
    def get_teacher_schedule(self, teacher_name):
        """
        获取指定教师的课程安排
        
        参数:
            teacher_name (str): 教师姓名
            
        返回:
            list: 该教师的所有课程记录列表
        """
        return [item for item in self.schedule_data if teacher_name in item['teacher']]
    
    def get_teacher_grades(self, teacher_name):
        """
        获取教师教授的年级统计
        
        参数:
            teacher_name (str): 教师姓名
            
        返回:
            dict: 年级 -> 课程数量的映射
        """
        schedule = self.get_teacher_schedule(teacher_name)
        grades = {}
        for item in schedule:
            grade = item['grade']
            if grade:
                grades[grade] = grades.get(grade, 0) + 1
        return grades
    
    def export_to_pdf(self, teacher_name, output_file):
        """
        导出教师课程表为PDF文件
        
        生成格式化的PDF课程表，包含：
        - 标题：教师姓名
        - 按星期分组的课程表
        - 表格列：时间、课程、年级、教室、教师
        
        参数:
            teacher_name (str): 教师姓名
            output_file (str): 输出PDF文件路径
            
        返回:
            bool: 成功返回True，失败返回False
        """
        # 获取该教师的课程安排
        schedule = self.get_teacher_schedule(teacher_name)
        
        if not schedule:
            print(f"Преподаватель {teacher_name} не найден")  # 教师未找到
            return False
        
        # 创建PDF文档
        # A4纸张，左右边距1.5cm，上下边距2cm
        doc = SimpleDocTemplate(
            output_file, 
            pagesize=A4, 
            rightMargin=1.5*cm, 
            leftMargin=1.5*cm, 
            topMargin=2*cm, 
            bottomMargin=2*cm
        )
        
        # 文档内容列表
        story = []
        # 获取示例样式
        styles = getSampleStyleSheet()
        
        # 定义标题样式
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Title'],
            fontName='Arial-Bold',
            fontSize=18,
            spaceAfter=20
        )
        
        # 定义星期标题样式
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontName='Arial-Bold',
            fontSize=14,
            spaceAfter=10
        )
        
        # 定义单元格内容样式（支持自动换行）
        cell_style = ParagraphStyle(
            'CellStyle',
            fontName='Arial',
            fontSize=8,       # 字体大小
            leading=10,       # 行高
            wordWrap='CJK'    # 支持中日韩文字换行
        )
        
        # 定义表头样式
        header_style = ParagraphStyle(
            'HeaderStyle',
            fontName='Arial-Bold',
            fontSize=9,
            leading=11,
            textColor=colors.whitesmoke  # 白色文字
        )
        
        # 添加文档标题
        title = Paragraph(f"Расписание преподавателя: {teacher_name}", title_style)
        story.append(title)
        story.append(Spacer(1, 0.5*cm))  # 添加间距
        
        # 星期顺序（俄语）
        days_order = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота']
        
        # 按星期生成课程表
        for day in days_order:
            # 筛选当天的课程
            day_schedule = [item for item in schedule if item['day'] == day]
            
            if day_schedule:
                # 添加星期标题
                day_title = Paragraph(f"<b>{day}</b>", heading_style)
                story.append(day_title)
                story.append(Spacer(1, 0.2*cm))
                
                # 构建表格数据
                # 表头行
                table_data = [[
                    Paragraph('<b>Время</b>', header_style),      # 时间
                    Paragraph('<b>Предмет</b>', header_style),    # 课程
                    Paragraph('<b>Курс</b>', header_style),       # 年级
                    Paragraph('<b>Аудитория</b>', header_style),  # 教室
                    Paragraph('<b>Преподаватель</b>', header_style)  # 教师
                ]]
                
                # 数据行（按时间排序）
                for item in sorted(day_schedule, key=lambda x: x['time']):
                    table_data.append([
                        Paragraph(item['time'], cell_style),
                        Paragraph(item['lesson'], cell_style),
                        Paragraph(str(item['grade']) if item['grade'] else '', cell_style),
                        Paragraph(item['room'], cell_style),
                        Paragraph(teacher_name, cell_style)
                    ])
                
                # 创建表格
                # 列宽：时间2.2cm, 课程8cm, 年级2.5cm, 教室2cm, 教师3.3cm
                table = Table(table_data, colWidths=[2.2*cm, 8*cm, 2.5*cm, 2*cm, 3.3*cm])
                
                # 设置表格样式
                table.setStyle(TableStyle([
                    # 表头背景色（灰色）
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    # 左对齐
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    # 顶部对齐（多行内容更美观）
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    # 内边距
                    ('TOPPADDING', (0, 0), (-1, -1), 4),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                    ('LEFTPADDING', (0, 0), (-1, -1), 4),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                    # 数据行背景色（米色）
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    # 网格线
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ]))
                
                story.append(table)
                story.append(Spacer(1, 0.5*cm))  # 添加间距
        
        # 生成PDF文件
        doc.build(story)
        print(f"PDF-файл создан: {output_file}")  # PDF文件已创建
        return True


def main():
    """
    模块独立运行时的主函数
    
    用于直接测试课程表处理器功能。
    """
    excel_file = '22ITsTiM_Raspisanie_2_polugodie_25-26_mag__pechat2.xlsx'
    
    # 创建处理器并加载数据
    processor = ScheduleProcessor(excel_file)
    processor.load_data()
    processor.parse_schedule()
    
    # 显示教师列表
    teachers = processor.get_teachers()
    print("\n可用的教师列表:")
    for i, teacher in enumerate(teachers, 1):
        print(f"{i}. {teacher}")
    
    # 获取用户选择
    print("\n请输入教师姓名或编号:")
    choice = input("> ").strip()
    
    # 解析选择
    if choice.isdigit() and 1 <= int(choice) <= len(teachers):
        selected_teacher = teachers[int(choice) - 1]
    else:
        selected_teacher = choice
    
    print(f"\n已选择: {selected_teacher}")
    
    # 生成PDF
    output_file = f"schedule_{selected_teacher.replace(' ', '_')}.pdf"
    if processor.export_to_pdf(selected_teacher, output_file):
        print(f"成功生成PDF文件: {output_file}")


# 模块入口点
if __name__ == "__main__":
    main()
