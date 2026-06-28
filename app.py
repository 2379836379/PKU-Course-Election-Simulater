"""
Virtual Course Selection Application

Expected columns in courses.xlsx:
- 课程号 (Course ID)
- 班号 (Class ID)
- 院系 (Department)
- 课程名 (Course Name)
- 参考学分 (Credits)
- 授课教师 (Instructor)
- 上课时间 (Time)
- 修读对象 (Target Audience)

Additional columns that may exist:
- 学年学期 (Academic Year/Term)
- 表格类型 (Table Type)
- 内部学期 (Internal Term)
- 课程英文名 (English Course Name)
- 课程类别 (Course Category)
- 周学时 (Weekly Hours)
- 总学时 (Total Hours)
- 起止周 (Start-End Weeks)
- 备注 (Notes)
"""

import streamlit as st
import pandas as pd
import numpy as np
from collections import defaultdict
import io
import base64
import json
import hashlib
import gc

# Enable Copy-on-Write for Pandas (Memory Optimization)
try:
    pd.options.mode.copy_on_write = True
except ImportError:
    pass  # Pandas version might be too old

# Language dictionary for internationalization
LANGUAGES = {
    "en": {
        "app_title": "Mock Course Selection Application",
        "department": "Department",
        "course_name": "Course Name",
        "class_id": "Class ID",
        "credits": "Credits",
        "instructor": "Instructor",
        "time": "Time",
        "select": "Select",
        "cancel": "Move to Pool",
        "user_department": "User Department",
        "second_department": "Second Department",
        "degree_type": "Degree Type",
        "single_degree": "Single Degree (Max 25 Credits)",
        "double_degree": "Double Degree (Max 30 Credits)",
        "filter_by_department": "Filter by Course Department",
        "filter_by_weekday": "Filter by Weekday",
        "filter_time_start": "Start Period",
        "filter_time_end": "End Period",
        "search_course": "Search by Course Name",
        "timetable": "Timetable",
        "current_credits": "Current Credits",
        "max_credits": "Max Credits",
        "warning": "Warning",
        "credit_exceeded": "Credit limit exceeded!",
        "conflict_detected": "Time conflict detected with",
        "no_conflict": "No conflicts. Course added successfully.",
        "page": "Page",
        "of": "of",
        "courses_per_page": "courses per page",
        "week_mon": "Mon",
        "week_tue": "Tue",
        "week_wed": "Wed",
        "week_thu": "Thu",
        "week_fri": "Fri",
        "week_sat": "Sat",
        "week_sun": "Sun",
        "periods": "Periods",
        "language": "Language",
        "chinese": "Chinese",
        "english": "English",
        "file_not_found": "courses.xlsx file not found. Please upload a file or generate sample data.",
        "upload_file": "Upload Excel File",
        "generate_sample": "Generate Sample Data",
        "all_departments": "All Departments",
        "all_courses": "All Courses",
        "selected_courses": "Selected Courses",
        "export_timetable": "Export Timetable",
        "export_success": "Timetable exported successfully!",
        "backup_restore": "Backup / Restore",
        "export_selected_json": "Export Selected (JSON)",
        "import_selected_json": "Import Selected (JSON)",
        "import_selected_button": "Restore Selection",
        "import_invalid_json": "Invalid JSON file",
        "import_done": "Selection restored",
        "import_partial": "Some courses were not found",
        "nav": "Navigation",
        "nav_courses": "Browse Courses",
        "nav_preselect": "Preselection + Timetable",
        "nav_timetable": "Timetable",
        "preselected_courses": "Preselected Courses",
        "add_to_preselect": "Preselect",
        "add_to_timetable": "Add to Timetable",
        "remove": "Remove from Pool",
        "already_selected": "Already selected",
        "already_in_pool": "Already in pool",
        "added_to_pool": "Added to preselection",
        "moved_to_timetable": "Added to timetable"
    },
    "zh": {
        "app_title": "模拟选课系统",
        "department": "院系",
        "course_name": "课程名",
        "class_id": "班号",
        "credits": "学分",
        "instructor": "授课教师",
        "time": "上课时间",
        "select": "选课",
        "cancel": "移回预选池",
        "user_department": "用户所在院系",
        "second_department": "第二学位院系",
        "degree_type": "学位类型",
        "single_degree": "单学位（最多25学分）",
        "double_degree": "双学位（最多30学分）",
        "filter_by_department": "按院系筛选",
        "filter_by_weekday": "按星期筛选",
        "filter_time_start": "起始节次",
        "filter_time_end": "结束节次",
        "search_course": "搜索课程名",
        "timetable": "课程表",
        "current_credits": "当前学分",
        "max_credits": "最大学分",
        "warning": "警告",
        "credit_exceeded": "超过学分限制！",
        "conflict_detected": "检测到时间冲突，与以下课程冲突：",
        "no_conflict": "无冲突。成功添加课程。",
        "page": "页码",
        "of": "页，共",
        "courses_per_page": "门课程每页",
        "week_mon": "周一",
        "week_tue": "周二",
        "week_wed": "周三",
        "week_thu": "周四",
        "week_fri": "周五",
        "week_sat": "周六",
        "week_sun": "周日",
        "periods": "节次",
        "language": "语言",
        "chinese": "中文",
        "english": "英文",
        "file_not_found": "未找到 courses.xlsx 文件。请上传文件或生成示例数据。",
        "upload_file": "上传Excel文件",
        "generate_sample": "生成示例数据",
        "all_departments": "所有院系",
        "all_courses": "所有课程",
        "selected_courses": "已选课程",
        "export_timetable": "导出课程表",
        "export_success": "课程表导出成功！",
        "backup_restore": "备份 / 恢复",
        "export_selected_json": "导出已选课程（JSON）",
        "import_selected_json": "导入已选课程（JSON）",
        "import_selected_button": "恢复选课",
        "import_invalid_json": "JSON 文件不合法",
        "import_done": "已恢复选课",
        "import_partial": "部分课程未找到",
        "nav": "页面",
        "nav_courses": "选课",
        "nav_preselect": "预选池/课表",
        "nav_timetable": "课表",
        "preselected_courses": "预选课程",
        "add_to_preselect": "加入预选",
        "add_to_timetable": "加入课表",
        "remove": "从预选池移除",
        "already_selected": "已在课表",
        "already_in_pool": "已在预选池",
        "added_to_pool": "已加入预选池",
        "moved_to_timetable": "已加入课表"
    }
}

@st.cache_resource(ttl=3600)  # Use cache_resource for shared read-only data (saves RAM)
def load_data():
    """Load course data from Parquet file with caching"""
    df = None
    loaded_from_parquet = False
    
    # Try Parquet first (Only source)
    try:
        df = pd.read_parquet("courses.parquet")
        loaded_from_parquet = True
    except FileNotFoundError:
        pass
    
    if df is None:
        return None
    
    # Process the data according to requirements
    # 1. Convert columns to category for memory optimization
    categorical_cols = ['院系', '班号', '课程类别', '学年学期', '表格类型', '内部学期']
    for col in categorical_cols:
        if col in df.columns:
            df[col] = df[col].astype('category')
            
    # 2. Optimize numeric types
    # Credits can be integers if no .5 exists, but usually credits can be 0.5, 1.5 etc.
    # User requested integers, but we must be safe. Let's check if we can downcast.
    # If all are integers, we use int8/int16. Otherwise float32.
    
    numeric_cols = ['参考学分', '周学时', '总学时']
    for col in numeric_cols:
        if col in df.columns:
            # First coerce to numeric
            s = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            # Check if all values are effectively integers (e.g. 2.0, 3.0)
            is_integer = np.all(np.mod(s, 1) == 0)
            
            if is_integer:
                # Safe to convert to integer (int16 is sufficient for credits)
                df[col] = s.astype('int16')
            else:
                # Must keep as float, but float32 is enough
                df[col] = s.astype('float32')

    # Preprocess helper columns
    # We NO LONGER pre-compute '_parsed_time' to save significant RAM.
    # It will be computed on-demand when a course is selected.

    # Lowercase helper for search (kept for speed)
    if '_course_name_lower' not in df.columns:
        df['_course_name_lower'] = df['课程名'].astype(str).str.lower()
    if '_instructor_lower' not in df.columns:
        if '授课教师' in df.columns:
            df['_instructor_lower'] = df['授课教师'].astype(str).str.lower()
        else:
            df['_instructor_lower'] = ""
    
    # Auto-save logic removed as we don't want to save derived columns or modify the source
    # unless it's strictly necessary. The original requirement was "keep import entry parquet only".
    # Since we are optimizing RAM, we don't want to bloat the parquet with _parsed_time either.
    
    # Force garbage collection after heavy loading
    gc.collect()
    
    return df

def generate_sample_data():
    """Generate sample course data for demonstration"""
    sample_data = {
        '课程号': ['CS101', 'CS102', 'CS201', 'CS202', 'CS301', 'CS302', 'CS401', 'CS402'],
        '班号': ['01', '01', '01', '01', '01', '01', '01', '01'],
        '院系': ['计算机学院', '计算机学院', '计算机学院', '计算机学院', '计算机学院', '计算机学院', '计算机学院', '计算机学院'],
        '课程名': ['计算机基础', 'Python编程', '数据结构', '算法分析', '数据库原理', '操作系统', '计算机网络', '软件工程'],
        '参考学分': [3, 3, 4, 4, 3, 3, 3, 3],
        '授课教师': ['张老师', '李老师', '王老师', '赵老师', '孙老师', '周老师', '吴老师', '郑老师'],
        '上课时间': [
            '周一1-2，周三3-4',
            '周二1-2单，周四3-4单',
            '周一3-4双，周五1-2双',
            '周二5-6，周四5-6',
            '周三7-8，周五3-4',
            '周一7-8单，周三7-8单',
            '周二7-8双，周四7-8双',
            '周五5-6'
        ],
        '修读对象': [
            '计算机学院学生',
            '全校学生在籍',
            '计算机学院学生',
            '计算机学院学生',
            '计算机学院学生',
            '计算机学院学生',
            '计算机学院学生',
            '计算机学院学生'
        ]
    }
    return pd.DataFrame(sample_data)


def parse_time(time_str):
    """
    Parse time string into structured data
    Format examples:
    - "周一1-2" (Every week)
    - "周二1-2单" (Odd weeks only)
    - "周三1-2双" (Even weeks only)
    - "周一1-2，周三3-4" (Multiple time slots)
    """
    if pd.isna(time_str):
        return []
    
    slots = str(time_str).split('，')
    parsed_slots = []
    
    for slot in slots:
        # Extract day, periods, and week type
        if '单' in slot:
            week_type = 'odd'
            slot = slot.replace('单', '')
        elif '双' in slot:
            week_type = 'even'
            slot = slot.replace('双', '')
        else:
            week_type = 'all'
        
        # Extract day and periods
        day_map = {
            '周一': 'mon', '周二': 'tue', '周三': 'wed',
            '周四': 'thu', '周五': 'fri', '周六': 'sat', '周日': 'sun'
        }
        
        day = None
        for chinese_day, english_day in day_map.items():
            if slot.startswith(chinese_day):
                day = english_day
                slot = slot[len(chinese_day):]
                break
        
        if day and '-' in slot:
            try:
                start_period, end_period = map(int, slot.split('-'))
                parsed_slots.append({
                    'day': day,
                    'start_period': start_period,
                    'end_period': end_period,
                    'week_type': week_type
                })
            except ValueError:
                continue  # Skip malformed entries
    
    return parsed_slots


def course_matches_day_time(time_str, selected_days, start_period, end_period):
    time_slots = parse_time(time_str)
    if not time_slots:
        return False
    for slot in time_slots:
        if slot["day"] not in selected_days:
            continue
        if slot["end_period"] < start_period or slot["start_period"] > end_period:
            continue
        return True
    return False


def export_selected_courses_json(selected_courses):
    items = []
    seen = set()
    for c in selected_courses:
        course_id = str(c.get("课程号", "")).strip()
        class_id = str(c.get("班号", "")).strip()
        if not course_id or not class_id:
            continue
        key = (course_id, class_id)
        if key in seen:
            continue
        seen.add(key)
        items.append(
            {
                "课程号": course_id,
                "班号": class_id,
                "课程名": str(c.get("课程名", "")).strip(),
                "授课教师": str(c.get("授课教师", "")).strip(),
            }
        )
    payload = {"version": 1, "courses": items}
    return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")


def import_selected_courses_json(df, raw_bytes):
    try:
        payload = json.loads(raw_bytes.decode("utf-8"))
    except Exception:
        return None, None

    if isinstance(payload, list):
        courses_list = payload
    elif isinstance(payload, dict) and isinstance(payload.get("courses"), list):
        courses_list = payload["courses"]
    else:
        return None, None

    if not {"课程号", "班号"}.issubset(set(df.columns)):
        return [], [{"课程号": x.get("课程号"), "班号": x.get("班号")} for x in courses_list if isinstance(x, dict)]

    index_df = df.set_index(["课程号", "班号"], drop=False)
    restored = []
    missing = []
    seen = set()

    for item in courses_list:
        if not isinstance(item, dict):
            continue
        course_id = str(item.get("课程号", "")).strip()
        class_id = str(item.get("班号", "")).strip()
        if not course_id or not class_id:
            continue
        key = (course_id, class_id)
        if key in seen:
            continue
        seen.add(key)

        try:
            row = index_df.loc[key]
            if isinstance(row, pd.DataFrame):
                row = row.iloc[0]
            course_dict = row.to_dict()
            course_dict["_parsed_time"] = parse_time(course_dict.get("上课时间"))
            restored.append(course_dict)
        except Exception:
            missing.append({"课程号": course_id, "班号": class_id})

    return restored, missing

def check_conflict(new_course, selected_courses):
    """
    Check if there's a time conflict between new course and selected courses
    Optimized with early exit strategy
    """
    if hasattr(new_course, 'to_dict'):
        new_course = new_course.to_dict()
    new_time_slots = (new_course.get('_parsed_time')
                      if isinstance(new_course, dict) else None)
    if not new_time_slots:
        new_time_slots = parse_time(new_course['上课时间'])
    
    # Early exit if no time slots
    if not new_time_slots:
        return None
    
    for course in selected_courses:
        existing_time_slots = course.get('_parsed_time') or parse_time(course['上课时间'])
        
        # Early exit if no existing time slots
        if not existing_time_slots:
            continue
        
        for new_slot in new_time_slots:
            for existing_slot in existing_time_slots:
                # Check if same day (early exit if not)
                if new_slot['day'] != existing_slot['day']:
                    continue
                
                # Check if periods overlap (early exit if not)
                if not (new_slot['start_period'] <= existing_slot['end_period'] and 
                        new_slot['end_period'] >= existing_slot['start_period']):
                    continue
                
                # Check if weeks overlap
                # No conflict if one is odd and the other is even
                if (new_slot['week_type'] == 'odd' and existing_slot['week_type'] == 'even') or \
                   (new_slot['week_type'] == 'even' and existing_slot['week_type'] == 'odd'):
                    continue
                
                # Conflict detected - return immediately
                return course['课程名']
    
    return None  # No conflict

def create_timetable(selected_courses, lang):
    """Create timetable visualization (optimized with cached time parsing)"""
    # Initialize timetable matrix (7 days x 12 periods) - Mon-Sun as requested
    days = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']
    day_names = {
        'mon': lang["week_mon"], 'tue': lang["week_tue"], 'wed': lang["week_wed"],
        'thu': lang["week_thu"], 'fri': lang["week_fri"], 'sat': lang["week_sat"], 
        'sun': lang["week_sun"]
    }
    
    timetable = {}
    for day in days:
        timetable[day] = {period: [] for period in range(1, 13)}
    
    # Fill timetable with selected courses
    for course in selected_courses:
        # Use cached parsed time if available
        if isinstance(course, dict) and '_parsed_time' in course:
            time_slots = course['_parsed_time']
        else:
            time_slots = parse_time(course['上课时间'])
        
        for slot in time_slots:
            day = slot['day']
            week_type = slot['week_type']
            course_display = f"{course['课程名']} ({course['班号']})"
            
            for period in range(slot['start_period'], slot['end_period'] + 1):
                if week_type == 'odd':
                    timetable[day][period].append({'course': course_display, 'week': 'odd'})
                elif week_type == 'even':
                    timetable[day][period].append({'course': course_display, 'week': 'even'})
                else:  # all weeks
                    timetable[day][period].append({'course': course_display, 'week': 'all'})
    
    return timetable, day_names

def export_timetable_to_excel(selected_courses, lang):
    """Export timetable to Excel file"""
    if not selected_courses:
        return None
    
    # Create timetable data
    timetable, day_names = create_timetable(selected_courses, lang)
    
    # Always show periods 1-12
    periods = list(range(1, 13))
    
    # Create timetable DataFrame for export
    timetable_df = pd.DataFrame(index=periods, columns=list(day_names.values()))
    
    for day_key, day_name in day_names.items():
        for period in periods:
            courses_in_slot = timetable[day_key][period]
            if courses_in_slot:
                # Format display based on week type
                course_texts = []
                for c in courses_in_slot:
                    if c['week'] == 'odd':
                        course_texts.append(f"{c['course']} [单]")
                    elif c['week'] == 'even':
                        course_texts.append(f"{c['course']} [双]")
                    else:
                        course_texts.append(c['course'])
                timetable_df.loc[period, day_name] = '\n'.join(course_texts)
            else:
                timetable_df.loc[period, day_name] = ""
    
    # Convert to Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        timetable_df.to_excel(writer, sheet_name='课程表' if lang.get('language') == 'zh' else 'Timetable')
        
        # Also export selected courses as a separate sheet with all necessary information for import
        selected_df = pd.DataFrame(selected_courses)
        selected_df.to_excel(writer, sheet_name='已选课程' if lang.get('language') == 'zh' else 'Selected Courses', index=False)
    
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="模拟选课", layout="wide")

    # Global CSS tweaks
    st.markdown(
        """
        <style>
        div[data-baseweb="toast"] {
            width: auto !important;
            min-height: auto !important;
            height: auto !important;
            max-width: 80vw !important;
        }
        div[data-baseweb="toast"] div {
            white-space: pre-wrap !important;
            word-wrap: break-word !important;
            line-height: 1.5 !important;
            height: auto !important;
        }
        div[data-baseweb="toast"] > div:last-child {
            align-items: flex-start !important;
            padding-top: 8px !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # Initialize page state
    if "current_page" not in st.session_state:
        st.session_state.current_page = 1

    # Language selection
    language = st.sidebar.selectbox(
        "Language / 语言",
        options=["zh", "en"],
        format_func=lambda x: "中文" if x == "zh" else "English",
    )
    lang = LANGUAGES[language]

    st.title(lang["app_title"])

    # Sidebar: cache clear
    with st.sidebar:
        if st.button("🔄 " + ("清除缓存" if language == "zh" else "Clear Cache")):
            st.cache_resource.clear()
            st.session_state.current_page = 1
            st.success("✓ " + ("缓存已清除" if language == "zh" else "Cache cleared"))
            st.rerun()

    # Load data
    df = load_data()
    if df is None:
        st.warning(lang["file_not_found"].replace("courses.xlsx", "courses.parquet"))

        col1, col2 = st.columns(2)
        with col1:
            st.info("Please place 'courses.parquet' in the application directory.")

        with col2:
            if st.button(lang["generate_sample"]):
                df = generate_sample_data()
                try:
                    save_df = df.copy()
                    save_df.to_parquet("courses.parquet", compression="snappy", index=False)
                except Exception:
                    pass

        if df is None:
            st.stop()

    # Session state
    if "selected_courses" not in st.session_state:
        st.session_state.selected_courses = []
    if "preselected_courses" not in st.session_state:
        st.session_state.preselected_courses = []

    if "degree_type" not in st.session_state:
        st.session_state.degree_type = "single"

    if "timetable_cache" not in st.session_state:
        st.session_state.timetable_cache = None
    if "timetable_courses_hash" not in st.session_state:
        st.session_state.timetable_courses_hash = None

    def course_key(course: dict) -> tuple[str, str]:
        return (str(course.get("课程号", "")).strip(), str(course.get("班号", "")).strip())

    def list_has(course_list: list[dict], key: tuple[str, str]) -> bool:
        for c in course_list:
            if course_key(c) == key:
                return True
        return False

    def remove_from_list(course_list: list[dict], key: tuple[str, str]) -> None:
        for i, c in enumerate(list(course_list)):
            if course_key(c) == key:
                course_list.pop(i)
                return

    # Navigation
    nav = st.sidebar.radio(
        lang.get("nav", "Navigation"),
        options=["courses", "preselect"],
        format_func=lambda x: (lang["nav_courses"] if x == "courses" else lang["nav_preselect"]),
    )

    # Degree type (affects credit warning)
    degree_type = st.sidebar.radio(
        lang["degree_type"],
        options=["single", "double"],
        format_func=lambda x: lang["single_degree"] if x == "single" else lang["double_degree"],
    )
    st.session_state.degree_type = degree_type
    max_credits = 25 if degree_type == "single" else 30
    current_credits = sum(float(c.get("参考学分", 0) or 0) for c in st.session_state.selected_courses)

    # Sidebar quick stats
    st.sidebar.caption(f"{lang['current_credits']}: {current_credits} / {lang['max_credits']}: {max_credits}")
    st.sidebar.caption(f"Pool: {len(st.session_state.preselected_courses)}" if language == "en" else f"预选池: {len(st.session_state.preselected_courses)}")
    st.sidebar.caption(f"Selected: {len(st.session_state.selected_courses)}" if language == "en" else f"已选: {len(st.session_state.selected_courses)}")

    def add_to_preselect(row_or_dict) -> None:
        course_dict = row_or_dict.to_dict() if hasattr(row_or_dict, "to_dict") else dict(row_or_dict)
        key = course_key(course_dict)
        if not key[0] or not key[1]:
            return

        if list_has(st.session_state.selected_courses, key):
            st.toast(lang.get("already_selected", "Already selected"), icon="ℹ️")
            return
        if list_has(st.session_state.preselected_courses, key):
            st.toast(lang.get("already_in_pool", "Already in pool"), icon="ℹ️")
            return

        st.session_state.preselected_courses.append(course_dict)
        st.toast(lang.get("added_to_pool", "Added to preselection"), icon="✅")

    def move_to_timetable(course_dict: dict) -> None:
        key = course_key(course_dict)
        if not key[0] or not key[1]:
            return

        if list_has(st.session_state.selected_courses, key):
            st.toast(lang.get("already_selected", "Already selected"), icon="ℹ️")
            remove_from_list(st.session_state.preselected_courses, key)
            st.rerun()

        conflict_course = check_conflict(course_dict, st.session_state.selected_courses)
        if conflict_course:
            st.toast(f"❌ {lang['conflict_detected']} {conflict_course}", icon="⚠️")
            return

        course_dict = dict(course_dict)
        course_dict["_parsed_time"] = parse_time(course_dict.get("上课时间"))
        st.session_state.selected_courses.append(course_dict)
        remove_from_list(st.session_state.preselected_courses, key)

        st.session_state.timetable_cache = None
        st.session_state.timetable_courses_hash = None
        st.toast(lang.get("moved_to_timetable", "Added to timetable"), icon="✅")
        st.rerun()

    def cancel_to_preselect(key: tuple[str, str]) -> None:
        # Move a course from timetable (selected) back to the preselection pool.
        st.session_state.cancel_clicks = int(st.session_state.get("cancel_clicks", 0)) + 1

        before_pool = len(st.session_state.preselected_courses)
        before_sel = len(st.session_state.selected_courses)

        removed = None
        for i, c in enumerate(list(st.session_state.selected_courses)):
            if course_key(c) == key:
                removed = dict(c)
                removed.pop("_parsed_time", None)
                st.session_state.selected_courses.pop(i)
                break

        if removed is None:
            st.session_state.last_cancel_debug = {
                "key": [key[0], key[1]],
                "status": "not_found_in_selected",
                "before_pool": before_pool,
                "after_pool": len(st.session_state.preselected_courses),
                "before_selected": before_sel,
                "after_selected": len(st.session_state.selected_courses),
                "cancel_clicks": st.session_state.cancel_clicks,
            }
            return

        remove_from_list(st.session_state.preselected_courses, key)
        st.session_state.preselected_courses.insert(0, removed)

        st.session_state.timetable_cache = None
        st.session_state.timetable_courses_hash = None
        st.session_state.last_moved_back = removed
        st.session_state.last_cancel_debug = {
            "key": [key[0], key[1]],
            "status": "moved",
            "before_pool": before_pool,
            "after_pool": len(st.session_state.preselected_courses),
            "before_selected": before_sel,
            "after_selected": len(st.session_state.selected_courses),
            "cancel_clicks": st.session_state.cancel_clicks,
        }


    def render_courses_view() -> None:
        # Filters only make sense on this page
        filtered_df = df

        st.sidebar.header("Filters")

        def reset_page_callback():
            st.session_state.current_page = 1

        all_depts = [lang["all_departments"]] + sorted(filtered_df["院系"].unique())
        dept_filter = st.sidebar.selectbox(
            lang["filter_by_department"],
            options=all_depts,
            on_change=reset_page_callback,
            key="dept_filter",
        )
        if dept_filter != lang["all_departments"]:
            filtered_df = filtered_df[filtered_df["院系"] == dept_filter]

        days_list = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
        weekday_filter = st.sidebar.multiselect(
            lang["filter_by_weekday"],
            options=days_list,
            default=days_list,
            format_func=lambda d: lang["week_" + d],
            on_change=reset_page_callback,
            key="weekday_filter",
        )

        start_period = int(
            st.sidebar.number_input(
                lang["filter_time_start"],
                min_value=1,
                max_value=12,
                value=1,
                step=1,
                on_change=reset_page_callback,
                key="start_period",
            )
        )
        end_period = int(
            st.sidebar.number_input(
                lang["filter_time_end"],
                min_value=1,
                max_value=12,
                value=12,
                step=1,
                on_change=reset_page_callback,
                key="end_period",
            )
        )
        if start_period > end_period:
            start_period, end_period = end_period, start_period

        # Day/time filter
        if weekday_filter:
            filtered_df = filtered_df[
                filtered_df["上课时间"].apply(
                    lambda s: course_matches_day_time(s, weekday_filter, start_period, end_period)
                )
            ]

        # Search
        query = st.sidebar.text_input(lang["search_course"], value="", on_change=reset_page_callback, key="search_query")
        q = (query or "").strip().lower()
        if q:
            keywords = [k for k in q.split() if k]
            if keywords:
                course_name_s = filtered_df.get("_course_name_lower", filtered_df["课程名"].astype(str).str.lower())
                instructor_s = filtered_df.get("_instructor_lower", filtered_df.get("授课教师", "").astype(str).str.lower())
                mask = pd.Series(True, index=filtered_df.index)
                for k in keywords:
                    mask &= course_name_s.str.contains(k, na=False, regex=False) | instructor_s.str.contains(k, na=False, regex=False)
                filtered_df = filtered_df[mask]

        # Pagination
        courses_per_page = 10
        total_courses = len(filtered_df)
        total_pages = (total_courses - 1) // courses_per_page + 1 if total_courses > 0 else 1

        if st.session_state.current_page > total_pages:
            st.session_state.current_page = total_pages
        if st.session_state.current_page < 1:
            st.session_state.current_page = 1

        if total_pages > 1:
            def prev_page_callback():
                st.session_state.current_page -= 1

            def next_page_callback():
                st.session_state.current_page += 1

            c1, c2, c3, c4, c5 = st.columns([6, 1, 2, 1, 6], vertical_alignment="center")
            with c2:
                st.button("←", on_click=prev_page_callback, disabled=(st.session_state.current_page <= 1), key="prev_top", use_container_width=True)
            with c3:
                st.button(f"{st.session_state.current_page} / {total_pages}", disabled=True, key="page_display_top", use_container_width=True)
            with c4:
                st.button("→", on_click=next_page_callback, disabled=(st.session_state.current_page >= total_pages), key="next_top", use_container_width=True)

        start_idx = (st.session_state.current_page - 1) * courses_per_page
        end_idx = start_idx + courses_per_page
        page_courses = filtered_df.iloc[start_idx:end_idx]

        if page_courses.empty:
            st.info("No courses" if language == "en" else "没有符合条件的课程")
            return

        for idx, (_, row) in enumerate(page_courses.iterrows()):
            course_id = str(row.get("课程号", "")).strip()
            class_id = str(row.get("班号", "")).strip()
            key = (course_id, class_id)
            in_selected = list_has(st.session_state.selected_courses, key)
            in_pool = list_has(st.session_state.preselected_courses, key)

            with st.container(border=True):
                c_info, c_action = st.columns([4, 1], vertical_alignment="center")
                with c_info:
                    st.markdown(
                        f"**{row['课程名']}** <span style='color:grey; font-size:0.9em'>({row['课程号']})</span>",
                        unsafe_allow_html=True,
                    )
                    meta_text = f"教师：{row.get('授课教师','')}  |  院系：{row.get('院系','')}  |  学分：{row.get('参考学分','')}  |  时间：{row.get('上课时间','')}"
                    st.caption(meta_text)

                with c_action:
                    label = lang.get("add_to_preselect", lang["select"])
                    disabled = in_selected or in_pool
                    if st.button(label, key=f"pre_{start_idx+idx}_{course_id}_{class_id}", use_container_width=True, disabled=disabled):
                        add_to_preselect(row)
                        st.rerun()

                    if in_selected:
                        st.caption(lang.get("already_selected", "Already selected"))
                    elif in_pool:
                        st.caption(lang.get("already_in_pool", "Already in pool"))

    def render_preselect_view() -> None:
        # Timetable first (top of the page)
        st.subheader(lang["timetable"])

        moved_back = st.session_state.pop("last_moved_back", None)
        if moved_back:
            st.info(
                (
                    f"已将 {moved_back.get('课程名', '')} ({moved_back.get('课程号', '')}) 移回预选池"
                    if language == "zh"
                    else f"Moved {moved_back.get('课程名', '')} ({moved_back.get('课程号', '')}) back to preselection"
                )
            )

        err = st.session_state.pop("last_cancel_error", None)
        if err:
            st.error(f"Cancel callback error: {err}")

        dbg = st.session_state.get("last_cancel_debug")
        show_debug = st.checkbox("Show debug", value=False, key="show_debug")
        st.caption(f"cancel_clicks: {int(st.session_state.get('cancel_clicks', 0))}")
        st.caption(f"cancel_button_hits: {int(st.session_state.get('cancel_button_hits', 0))}")
        st.caption(f"last_cancel_button_key: {st.session_state.get('last_cancel_button_key')}")
        if show_debug:
            st.json(dbg or {"info": "No cancel debug captured yet."})
        st.session_state.render_seq = int(st.session_state.get('render_seq', 0)) + 1
        st.caption(f"render_seq: {int(st.session_state.get('render_seq', 0))}")

        if st.button("Debug: test button", key="debug_test_button"):
            st.session_state.debug_test_hits = int(st.session_state.get('debug_test_hits', 0)) + 1
        st.caption(f"debug_test_hits: {int(st.session_state.get('debug_test_hits', 0))}")


        timetable_header_bg = "var(--secondary-background-color)"
        timetable_header_fg = "var(--text-color)"

        day_names = {
            "mon": lang["week_mon"],
            "tue": lang["week_tue"],
            "wed": lang["week_wed"],
            "thu": lang["week_thu"],
            "fri": lang["week_fri"],
            "sat": lang["week_sat"],
            "sun": lang["week_sun"],
        }
        days_list = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]

        if st.session_state.selected_courses:
            courses_hash = hash(
                str([(c.get("课程号"), c.get("班号")) for c in st.session_state.selected_courses])
            )
            if (
                st.session_state.timetable_courses_hash == courses_hash
                and st.session_state.timetable_cache is not None
            ):
                timetable, _day_names = st.session_state.timetable_cache
            else:
                timetable, _day_names = create_timetable(st.session_state.selected_courses, lang)
                st.session_state.timetable_cache = (timetable, _day_names)
                st.session_state.timetable_courses_hash = courses_hash

            timetable_data_for_df = {}
            for day in days_list:
                day_col = []
                for period in range(1, 13):
                    courses = timetable[day][period]
                    if courses:
                        text_parts = []
                        for c in courses:
                            info = c["course"]
                            if c["week"] == "odd":
                                info += " (单)" if language == "zh" else " (Odd)"
                            elif c["week"] == "even":
                                info += " (双)" if language == "zh" else " (Even)"
                            text_parts.append(info)
                        day_col.append("\n".join(text_parts))
                    else:
                        day_col.append("")
                timetable_data_for_df[day_names[day]] = day_col

            df_tt = pd.DataFrame(timetable_data_for_df, index=range(1, 13))
        else:
            df_tt = pd.DataFrame(
                {day_names[d]: [""] * 12 for d in days_list}, index=range(1, 13)
            )

        def color_courses(val):
            if val and str(val).strip() != "":
                return "background-color: rgba(28, 131, 225, 0.2); border-radius: 4px; font-weight: bold; color: inherit;"
            return ""

        styled_df = (
            df_tt.style.map(color_courses)
            .set_properties(
                **{
                    "height": "65px",
                    "vertical-align": "middle",
                    "text-align": "center",
                    "white-space": "pre-wrap",
                    "border": "1px solid #444" if language == "zh" else "1px solid #ddd",
                }
            )
            .set_table_styles(
                [
                    {
                        "selector": "table",
                        "props": [
                            ("width", "100%"),
                            ("table-layout", "fixed"),
                            ("border-collapse", "collapse"),
                        ],
                    },
                    {
                        "selector": "th",
                        "props": [
                            ("background-color", timetable_header_bg),
                            ("color", timetable_header_fg),
                            ("text-align", "center"),
                            ("vertical-align", "middle"),
                        ],
                    },
                    {"selector": "td, th", "props": [("box-sizing", "border-box")]},
                ]
            )
        )

        st.markdown(
            f"""
<style>
.timetable-wrapper table {{
    width: 100% !important;
    table-layout: fixed !important;
    border-collapse: collapse !important;
}}
.timetable-wrapper th:first-child,
.timetable-wrapper td:first-child {{
    width: 6% !important;
}}
.timetable-wrapper th:not(:first-child),
.timetable-wrapper td:not(:first-child) {{
    width: 13.4% !important;
}}
.timetable-wrapper td, .timetable-wrapper th {{
    text-align: center !important;
    vertical-align: middle !important;
}}
</style>
<div class="timetable-wrapper">{styled_df.to_html()}</div>
""",
            unsafe_allow_html=True,
        )

        excel_data = export_timetable_to_excel(st.session_state.selected_courses, lang)
        if excel_data:
            st.download_button(
                label=lang["export_timetable"],
                data=excel_data,
                file_name="课程表.xlsx" if language == "zh" else "timetable.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.divider()

        # Preselection pool (below timetable)
        st.subheader(lang.get("preselected_courses", "Preselected Courses"))
        pool = st.session_state.preselected_courses
        if not pool:
            st.info(
                "预选池为空。请先在“选课”页将课程加入预选池。"
                if language == "zh"
                else "Preselection pool is empty. Add courses from the Browse page."
            )
        else:
            for idx, course in enumerate(list(pool)):
                key = course_key(course)
                with st.container(border=True):
                    c_info, c_action = st.columns([4, 2], vertical_alignment="center")
                    with c_info:
                        st.markdown(
                            f"**{course.get('课程名','')}** <span style='color:grey; font-size:0.9em'>({course.get('课程号','')})</span>",
                            unsafe_allow_html=True,
                        )
                        meta_text = (
                            f"教师：{course.get('授课教师','')}  |  院系：{course.get('院系','')}  |  学分：{course.get('参考学分','')}  |  时间：{course.get('上课时间','')}"
                        )
                        st.caption(meta_text)

                    with c_action:
                        b1, b2 = st.columns(2)
                        with b1:
                            if st.button(
                                lang.get("add_to_timetable", "Add to Timetable"),
                                key=f"to_tt_{idx}_{key[0]}_{key[1]}",
                                use_container_width=True,
                            ):
                                move_to_timetable(course)
                        with b2:
                            if st.button(
                                lang.get("remove", "Remove"),
                                key=f"rm_pre_{idx}_{key[0]}_{key[1]}",
                                use_container_width=True,
                                type="secondary",
                            ):
                                remove_from_list(st.session_state.preselected_courses, key)
                                st.rerun()

        st.divider()

        # Credits + backup/restore + selected list
        st.subheader(f"{lang['current_credits']}: {current_credits} / {lang['max_credits']}: {max_credits}")
        if current_credits > max_credits:
            st.markdown(
                f"""
                <div style="
                    background-color: rgba(255, 75, 75, 0.1);
                    color: rgb(163, 6, 6);
                    padding: 20px;
                    border-radius: 0.5rem;
                    border: 1px solid rgba(255, 75, 75, 0.2);
                    text-align: center;
                    margin-top: 10px;
                    margin-bottom: 10px;">
                    ⚠️ {lang['warning']}: {lang['credit_exceeded']}
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.subheader(lang["backup_restore"])
        selected_json = export_selected_courses_json(st.session_state.selected_courses)
        st.download_button(
            label=lang["export_selected_json"],
            data=selected_json,
            file_name="已选课程.json" if language == "zh" else "selected_courses.json",
            mime="application/json",
        )
        uploaded_json = st.file_uploader(
            label=lang["import_selected_json"],
            type=["json"],
            key="selected_courses_json_uploader",
        )
        if uploaded_json is not None:
            raw = uploaded_json.getvalue()
            sig = hashlib.sha256(raw).hexdigest()
            if st.session_state.get("selected_courses_import_sig") != sig:
                st.session_state.selected_courses_import_sig = sig
                restored, missing = import_selected_courses_json(df, raw)
                if restored is None:
                    st.toast(lang["import_invalid_json"], icon="⚠️")
                else:
                    st.session_state.selected_courses = restored
                    st.session_state.timetable_cache = None
                    st.session_state.timetable_courses_hash = None
                    st.toast(f"✅ {lang['import_done']}: {len(restored)}", icon="🎉")
                    if missing:
                        st.toast(f"⚠️ {lang['import_partial']}: {len(missing)}", icon="⚠️")
                    st.rerun()

        if st.session_state.selected_courses:
            st.subheader(lang["selected_courses"])
            for idx, course in enumerate(list(st.session_state.selected_courses)):
                key = course_key(course)
                with st.container(border=True):
                    c_info, c_action = st.columns([4, 1], vertical_alignment="center")
                    with c_info:
                        st.markdown(
                            f"**{course.get('课程名','')}** <span style='color:grey; font-size:0.9em'>({course.get('课程号','')})</span>",
                            unsafe_allow_html=True,
                        )
                        meta_text = (
                            f"教师：{course.get('授课教师','')}  |  院系：{course.get('院系','')}  |  学分：{course.get('参考学分','')}  |  时间：{course.get('上课时间','')}"
                        )
                        st.caption(meta_text)

                    with c_action:
                        if st.button(
                            lang["cancel"],
                            key=f"cancel_{idx}_{key[0]}_{key[1]}",
                            use_container_width=True,
                            type="primary",
                        ):
                            st.session_state.cancel_button_hits = int(
                                st.session_state.get("cancel_button_hits", 0)
                            ) + 1
                            st.session_state.last_cancel_button_key = [key[0], key[1]]
                            cancel_to_preselect(key)
                            st.rerun()


    def render_timetable_view() -> None:
        st.subheader(lang["timetable"])

        timetable_header_bg = "var(--secondary-background-color)"
        timetable_header_fg = "var(--text-color)"

        day_names = {
            'mon': lang["week_mon"], 'tue': lang["week_tue"], 'wed': lang["week_wed"],
            'thu': lang["week_thu"], 'fri': lang["week_fri"], 'sat': lang["week_sat"],
            'sun': lang["week_sun"],
        }
        days_list = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']

        if st.session_state.selected_courses:
            courses_hash = hash(str([(c.get('课程号'), c.get('班号')) for c in st.session_state.selected_courses]))
            if (st.session_state.timetable_courses_hash == courses_hash and st.session_state.timetable_cache is not None):
                timetable, _day_names = st.session_state.timetable_cache
            else:
                timetable, _day_names = create_timetable(st.session_state.selected_courses, lang)
                st.session_state.timetable_cache = (timetable, _day_names)
                st.session_state.timetable_courses_hash = courses_hash

            timetable_data_for_df = {}
            for day in days_list:
                day_col = []
                for period in range(1, 13):
                    courses = timetable[day][period]
                    if courses:
                        text_parts = []
                        for c in courses:
                            info = c['course']
                            if c['week'] == 'odd':
                                info += " (单)" if language == 'zh' else " (Odd)"
                            elif c['week'] == 'even':
                                info += " (双)" if language == 'zh' else " (Even)"
                            text_parts.append(info)
                        day_col.append("\n".join(text_parts))
                    else:
                        day_col.append("")
                timetable_data_for_df[day_names[day]] = day_col

            df_tt = pd.DataFrame(timetable_data_for_df, index=range(1, 13))
        else:
            df_tt = pd.DataFrame({day_names[d]: [""] * 12 for d in days_list}, index=range(1, 13))

        def color_courses(val):
            if val and str(val).strip() != "":
                return 'background-color: rgba(28, 131, 225, 0.2); border-radius: 4px; font-weight: bold; color: inherit;'
            return ''

        styled_df = (
            df_tt.style.map(color_courses)
            .set_properties(
                **{
                    'height': '65px',
                    'vertical-align': 'middle',
                    'text-align': 'center',
                    'white-space': 'pre-wrap',
                    'border': '1px solid #444' if language == 'zh' else '1px solid #ddd',
                }
            )
            .set_table_styles(
                [
                    {'selector': 'table', 'props': [('width', '100%'), ('table-layout', 'fixed'), ('border-collapse', 'collapse')]},
                    {'selector': 'th', 'props': [('background-color', timetable_header_bg), ('color', timetable_header_fg), ('text-align', 'center'), ('vertical-align', 'middle')]},
                    {'selector': 'td, th', 'props': [('box-sizing', 'border-box')]},
                ]
            )
        )

        st.markdown(
            f"""
<style>
.timetable-wrapper table {{
    width: 100% !important;
    table-layout: fixed !important;
    border-collapse: collapse !important;
}}
.timetable-wrapper th:first-child,
.timetable-wrapper td:first-child {{
    width: 6% !important;
}}
.timetable-wrapper th:not(:first-child),
.timetable-wrapper td:not(:first-child) {{
    width: 13.4% !important;
}}
.timetable-wrapper td, .timetable-wrapper th {{
    text-align: center !important;
    vertical-align: middle !important;
}}
</style>
<div class="timetable-wrapper">{styled_df.to_html()}</div>
""",
            unsafe_allow_html=True,
        )

        # Export timetable
        excel_data = export_timetable_to_excel(st.session_state.selected_courses, lang)
        if excel_data:
            st.download_button(
                label=lang["export_timetable"],
                data=excel_data,
                file_name="课程表.xlsx" if language == "zh" else "timetable.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        # Backup / Restore (selected courses)
        st.subheader(lang["backup_restore"])
        selected_json = export_selected_courses_json(st.session_state.selected_courses)
        st.download_button(
            label=lang["export_selected_json"],
            data=selected_json,
            file_name="已选课程.json" if language == "zh" else "selected_courses.json",
            mime="application/json",
        )
        uploaded_json = st.file_uploader(
            label=lang["import_selected_json"],
            type=["json"],
            key="selected_courses_json_uploader",
        )
        if uploaded_json is not None:
            raw = uploaded_json.getvalue()
            sig = hashlib.sha256(raw).hexdigest()
            if st.session_state.get("selected_courses_import_sig") != sig:
                st.session_state.selected_courses_import_sig = sig
                restored, missing = import_selected_courses_json(df, raw)
                if restored is None:
                    st.toast(lang["import_invalid_json"], icon="⚠️")
                else:
                    st.session_state.selected_courses = restored
                    st.session_state.timetable_cache = None
                    st.session_state.timetable_courses_hash = None
                    st.toast(f"✅ {lang['import_done']}: {len(restored)}", icon="🎉")
                    if missing:
                        st.toast(f"⚠️ {lang['import_partial']}: {len(missing)}", icon="⚠️")
                    st.rerun()

        # Credits
        st.subheader(f"{lang['current_credits']}: {current_credits} / {lang['max_credits']}: {max_credits}")
        if current_credits > max_credits:
            st.markdown(
                f"""
                <div style="
                    background-color: rgba(255, 75, 75, 0.1);
                    color: rgb(163, 6, 6);
                    padding: 20px;
                    border-radius: 0.5rem;
                    border: 1px solid rgba(255, 75, 75, 0.2);
                    text-align: center;
                    margin-top: 10px;
                    margin-bottom: 10px;">
                    ⚠️ {lang['warning']}: {lang['credit_exceeded']}
                </div>
                """,
                unsafe_allow_html=True,
            )
        # Selected courses list
        if st.session_state.selected_courses:
            st.subheader(lang["selected_courses"])
            for idx, course in enumerate(list(st.session_state.selected_courses)):
                key = course_key(course)
                with st.container(border=True):
                    c_info, c_action = st.columns([4, 1], vertical_alignment="center")

                    with c_info:
                        st.markdown(
                            f"**{course.get('课程名','')}** <span style='color:grey; font-size:0.9em'>({course.get('课程号','')})</span>",
                            unsafe_allow_html=True,
                        )
                        meta_text = (
                            f"教师：{course.get('授课教师','')}  |  院系：{course.get('院系','')}  |  学分：{course.get('参考学分','')}  |  时间：{course.get('上课时间','')}"
                        )
                        st.caption(meta_text)
                    with c_action:
                        if st.button(
                            lang["cancel"],
                            key=f"cancel_{idx}_{key[0]}_{key[1]}",
                            use_container_width=True,
                            type="primary",
                        ):
                            st.session_state.cancel_button_hits = int(st.session_state.get('cancel_button_hits', 0)) + 1
                            st.session_state.last_cancel_button_key = [key[0], key[1]]
                            cancel_to_preselect(key)
                            st.rerun()

    if nav == "courses":
        render_courses_view()
    else:
        render_preselect_view()


if __name__ == "__main__":
    main()









