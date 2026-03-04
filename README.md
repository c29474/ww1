# 课程表生成器

这是一个用于处理学院课程表的Python程序，可以从Excel文件中读取课程信息，并按教师筛选后导出为PDF格式。

**[Русская версия](README_RU.md)**

## 功能特点

- 从Excel文件读取课程表数据
- 自动解析教师信息（共59位教师，549条课程记录）
- 按教师筛选课程
- 生成分天显示的PDF课程表
- 支持搜索教师功能
- 友好的命令行界面

## 安装依赖

```bash
pip install -r requirements.txt
```

或手动安装：

```bash
pip install pandas openpyxl reportlab
```

## 使用方法

### 方式1: 使用交互式界面（推荐）

```bash
python main.py
```

程序会显示教师列表，您可以：
- 输入编号选择教师
- 使用 n/N 下一页，p/P 上一页浏览教师列表
- 使用 s/S 搜索教师（支持部分匹配）
- 使用 q/Q 退出程序

### 方式2: 在代码中使用

```python
from schedule_processor import ScheduleProcessor

# 创建处理器
processor = ScheduleProcessor('ITsTiM_Raspisanie_2_polugodie_25-26_bak__pechat.xlsx')

# 加载和解析数据
processor.load_data()
processor.parse_schedule()

# 获取所有教师
teachers = processor.get_teachers()

# 获取特定教师的课程
teacher_name = "Бурцев В."
schedule = processor.get_teacher_schedule(teacher_name)

# 导出为PDF
processor.export_to_pdf(teacher_name, "output.pdf")
```

## 文件说明

- `schedule_processor.py` - 核心处理类，包含数据加载、解析和PDF生成功能
- `main.py` - 交互式主程序
- `requirements.txt` - Python依赖包列表
- `ITsTiM_Raspisanie_2_polugodie_25-26_bak__pechat.xlsx` - 原始课程表文件

## 输出格式

生成的PDF文件包含：
- 教师姓名作为标题
- 按星期分组显示课程（Понедельник, Вторник, Среда, Четверг, Пятница, Суббота）
- 每节课显示：时间、课程名称、课程类型（лк/пз/лб）、教室

## 示例

选择教师 "Бурцев В." 后，程序会生成文件 `schedule_Бурцев_В_.pdf`，包含该教师的所有课程安排。

## 注意事项

- 确保Excel文件在同一目录下
- 教师姓名格式为 "姓 名."（如：Бурцев В.）
- PDF文件会保存在当前目录下
