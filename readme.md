# Mock Course Selection App (Streamlit)

A small Streamlit app for **mock course selection + timetable visualization**.

## Requirements
- Python 3
- Install dependencies: `pip install -r requirements.txt`

## Quick Start

It is recommended to run everything inside `PKU-Course-Election/` (the app reads `courses.parquet` via a relative path).

1. Prepare data (recommended): put `test.xlsx` under `PKU-Course-Election/` and convert it to `courses.parquet`:
   - `python .\\convert_data.py`
   - Or from the repo root: `python .\\PKU-Course-Election\\convert_data.py`
2. Start the app:
   - `cd .\\PKU-Course-Election`
   - `streamlit run app.py`

If `PKU-Course-Election/` does not contain `courses.parquet`, the app will warn you and provide a **Generate Sample Data** button (it will write `courses.parquet`).

## Data Prep (Excel → Parquet)

### Required Columns
By default, the converter reads `test.xlsx` from the **current working directory** (you can also pass `--excel`).
The Excel file must contain at least these columns:
- `课程号`
- `班号`
- `院系`
- `课程名`
- `参考学分`
- `授课教师`
- `上课时间`
- `修读对象`

### Optional Columns
If present, these columns are preserved as-is in `courses.parquet`:
- `学年学期`, `表格类型`, `内部学期`, `课程英文名`, `课程类别`, `周学时`, `总学时`, `起止周`, `备注`

### `上课时间` Format
The app parses this field for **conflict detection / timetable rendering / filtering**:
- Basic format: `周X起始-结束`
  - Example: `周一1-2`
- Odd/even weeks: append `单` or `双`
  - Example: `周二1-2单`, `周三3-4双`
- Multiple slots: separated by Chinese comma `，`
  - Example: `周一1-2，周三3-4`

## Converter (`convert_data.py`)
Running `python convert_data.py` will:
- Read `test.xlsx`
- Deduplicate by `课程号 + 班号`
  - If the same course appears in multiple rows, `修读对象` will be de-duplicated and joined by `，`
  - Other fields use the first row in the group
- Write `courses.parquet` (snappy compression)

## App Features (`app.py`)

### Browse / Filter / Search
Sidebar **Filters**:
- Filter by `院系`
- Filter by weekday (multi-select: Mon–Sun)
- Filter by period range (start/end periods, 1–12)

Search:
- Space-separated keywords: `kw1 kw2 ...`
- Matches within: `课程名` + `授课教师`
- Logic: every keyword must match, but can match either field

### Preselection Pool + Timetable
The app uses a two-step workflow:
- Canceling a course from the timetable will move it back to the preselection pool.
- On the **Browse Courses** page, click **Preselect** to add a course into the **preselection pool** (no conflict check at this stage).
- On the **Preselection + Timetable** page (timetable shown at the top), click **Add to Timetable** to run conflict detection and add it to the timetable.

Conflict rules:
- Same day + overlapping periods
- Week type overlaps (odd vs even is NOT a conflict; “every week” conflicts with anything)

### Credits
- Single degree: max 25
- Double degree: max 30
- Exceeding the limit shows a warning (does not block selection)

### Timetable Export
- Fixed timetable layout: Mon–Sun, periods 1–12
- Export to Excel:
  - Timetable sheet
  - Selected courses sheet (full fields)

### Cache / Performance
- Data loading is cached (sidebar has **Clear Cache**)
- To save memory, time parsing is computed on-demand and cached per selected course


