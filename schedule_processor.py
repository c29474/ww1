import pandas as pd
import re
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

class ScheduleProcessor:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.df = None
        self.teachers = set()
        self.schedule_data = []
        self.group_grade_map = {}
        self._register_fonts()
        
    def _register_fonts(self):
        try:
            font_path = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts', 'arial.ttf')
            if os.path.exists(font_path):
                pdfmetrics.registerFont(TTFont('Arial', font_path))
                pdfmetrics.registerFont(TTFont('Arial-Bold', font_path.replace('arial.ttf', 'arialbd.ttf')))
            else:
                font_path = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts', 'arial.ttf')
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont('Arial', font_path))
        except Exception as e:
            print(f"字体注册警告: {e}")
        
    def load_data(self):
        xl = pd.ExcelFile(self.excel_file)
        sheet_names = xl.sheet_names
        
        if 'Бак_2024-2025' in sheet_names:
            sheet_name = 'Бак_2024-2025'
        elif '表6' in sheet_names:
            sheet_name = '表6'
        else:
            sheet_name = sheet_names[0]
            print(f"Предупреждение: используется лист {sheet_name}")
        
        self.df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
        print(f"Успешно загружено расписание, {len(self.df)} строк")
        
        self._extract_grade_info()
    
    def _extract_grade_info(self):
        self.group_grade_map = {}
        if len(self.df) > 2:
            for col_idx in range(2, len(self.df.columns)):
                group_name = self.df.iloc[2, col_idx]
                if pd.notna(group_name) and isinstance(group_name, str):
                    grade = self._extract_grade_from_group(group_name)
                    self.group_grade_map[col_idx] = {
                        'group': group_name,
                        'grade': grade
                    }
        print(f"Извлечена информация о {len(self.group_grade_map)} группах")
    
    def _extract_grade_from_group(self, group_name):
        match = re.search(r'(\d+)', group_name)
        if match:
            return int(match.group(1))
        return ''
        
    def parse_schedule(self):
        self.schedule_data = []
        current_day = None
        current_time = None
        
        for idx, row in self.df.iterrows():
            day = row[0]
            time = row[1]
            
            if pd.notna(day) and 'П О Н Е Д Е Л Ь Н И К' in str(day):
                current_day = 'Понедельник'
            elif pd.notna(day) and 'В Т О Р Н И К' in str(day):
                current_day = 'Вторник'
            elif pd.notna(day) and 'С Р Е Д А' in str(day):
                current_day = 'Среда'
            elif pd.notna(day) and 'Ч Е Т В Е Р Г' in str(day):
                current_day = 'Четверг'
            elif pd.notna(day) and 'П Я Т Н И Ц А' in str(day):
                current_day = 'Пятница'
            elif pd.notna(day) and 'С У Б Б О Т А' in str(day):
                current_day = 'Суббота'
                
            if pd.notna(time):
                current_time = str(time).replace('\n', ' ')
            
            if current_day and current_time and idx > 0:
                for col_idx in range(2, len(row)):
                    cell_value = row[col_idx]
                    if pd.notna(cell_value) and isinstance(cell_value, str) and len(cell_value.strip()) > 5:
                        self._parse_cell(cell_value, current_day, current_time, col_idx)
        
        print(f"Парсинг завершен, найдено {len(self.schedule_data)} записей")
        print(f"Всего {len(self.teachers)} преподавателей")
        
    def _parse_cell(self, cell_value, day, time, col_idx):
        cell_value = str(cell_value).strip()
        
        teacher_match = re.search(r'([А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ]\.){1,2})', cell_value)
        if teacher_match:
            teacher = teacher_match.group(1)
            self.teachers.add(teacher)
            
            room_match = re.search(r'(\d+-\d+[а-я]?)|(\d+-\d+)', cell_value)
            room = room_match.group(0) if room_match else ''
            
            lesson_type_match = re.search(r'\((лк|пз|лб)\)', cell_value)
            lesson_type = lesson_type_match.group(1) if lesson_type_match else ''
            
            lesson_name = re.sub(r'\s*\([^)]*\)\s*', ' ', cell_value)
            lesson_name = re.sub(r'\s*[А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ]\.){1,2}\s*', ' ', lesson_name)
            lesson_name = re.sub(r'\s*\d+-\d+[а-я]?\s*', ' ', lesson_name)
            lesson_name = ' '.join(lesson_name.split())
            
            grade_info = self.group_grade_map.get(col_idx, {})
            grade = grade_info.get('grade', '')
            group_name = grade_info.get('group', '')
            
            if not grade:
                grade = ''
            
            self.schedule_data.append({
                'day': day,
                'time': time,
                'lesson': lesson_name,
                'type': lesson_type,
                'teacher': teacher,
                'room': room,
                'group_col': col_idx,
                'grade': str(grade) if grade else '',
                'group': group_name
            })
    
    def get_teachers(self):
        return sorted(list(self.teachers))
    
    def get_teacher_schedule(self, teacher_name):
        return [item for item in self.schedule_data if teacher_name in item['teacher']]
    
    def get_teacher_grades(self, teacher_name):
        schedule = self.get_teacher_schedule(teacher_name)
        grades = {}
        for item in schedule:
            grade = item['grade']
            if grade:
                grades[grade] = grades.get(grade, 0) + 1
        return grades
    
    def export_to_pdf(self, teacher_name, output_file):
        schedule = self.get_teacher_schedule(teacher_name)
        
        if not schedule:
            print(f"Преподаватель {teacher_name} не найден")
            return False
        
        doc = SimpleDocTemplate(output_file, pagesize=A4, rightMargin=2*cm, leftMargin=2*cm, 
                                topMargin=2*cm, bottomMargin=2*cm)
        story = []
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Title'],
            fontName='Arial-Bold',
            fontSize=18,
            spaceAfter=20
        )
        
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontName='Arial-Bold',
            fontSize=14,
            spaceAfter=10
        )
        
        title = Paragraph(f"Расписание преподавателя: {teacher_name}", title_style)
        story.append(title)
        story.append(Spacer(1, 0.5*cm))
        
        days_order = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота']
        
        for day in days_order:
            day_schedule = [item for item in schedule if item['day'] == day]
            if day_schedule:
                day_title = Paragraph(f"<b>{day}</b>", heading_style)
                story.append(day_title)
                story.append(Spacer(1, 0.2*cm))
                
                table_data = [['Время', 'Предмет', 'Тип', 'Курс', 'Аудитория']]
                for item in sorted(day_schedule, key=lambda x: x['time']):
                    table_data.append([
                        item['time'],
                        item['lesson'],
                        item['type'],
                        str(item['grade']) if item['grade'] else '',
                        item['room']
                    ])
                
                table = Table(table_data, colWidths=[2.5*cm, 9*cm, 2*cm, 1.5*cm, 2.5*cm], repeatRows=1)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                    ('FONTSIZE', (0, 0), (-1, -1), 9),
                    ('TOPPADDING', (0, 0), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                    ('LEFTPADDING', (0, 0), (-1, -1), 6),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))
                story.append(table)
                story.append(Spacer(1, 0.5*cm))
        
        grades = self.get_teacher_grades(teacher_name)
        if grades:
            story.append(Spacer(1, 0.3*cm))
            grades_title = Paragraph("<b>Распределение по курсам:</b>", heading_style)
            story.append(grades_title)
            
            grades_table_data = [['Курс', 'Количество занятий']]
            for grade in sorted(grades.keys(), key=int):
                grades_table_data.append([grade, str(grades[grade])])
            
            grades_table = Table(grades_table_data, colWidths=[3*cm, 4*cm])
            grades_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('TOPPADDING', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            story.append(grades_table)
            story.append(Spacer(1, 0.5*cm))
        
        doc.build(story)
        print(f"PDF-файл создан: {output_file}")
        return True

def main():
    excel_file = '22ITsTiM_Raspisanie_2_polugodie_25-26_mag__pechat2.xlsx'
    
    processor = ScheduleProcessor(excel_file)
    processor.load_data()
    processor.parse_schedule()
    
    teachers = processor.get_teachers()
    print("\n可用的教师列表:")
    for i, teacher in enumerate(teachers, 1):
        print(f"{i}. {teacher}")
    
    print("\n请输入教师姓名或编号:")
    choice = input("> ").strip()
    
    if choice.isdigit() and 1 <= int(choice) <= len(teachers):
        selected_teacher = teachers[int(choice) - 1]
    else:
        selected_teacher = choice
    
    print(f"\n已选择: {selected_teacher}")
    
    output_file = f"schedule_{selected_teacher.replace(' ', '_')}.pdf"
    if processor.export_to_pdf(selected_teacher, output_file):
        print(f"成功生成PDF文件: {output_file}")

if __name__ == "__main__":
    main()
