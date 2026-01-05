# Sommerferieplan

A Python script that calculates and optimizes summer vacation schedules for ATC (Air Traffic Control) employees.

## Overview

This tool helps plan summer vacation schedules by:
- **vacation_calculator.py**: Analyzes coverage requirements and calculates theoretical maximum vacation days
- **vacation_scheduler.py**: Optimizes vacation distribution across employees while maintaining shift coverage

The scheduler considers:
- Employee skills and qualifications
- Daily shift requirements by skill type  
- Minimum coverage requirements for each shift
- Weekday vs weekend coverage differences
- **Equal distribution** - all employees get the same vacation days (within 1-day tolerance)
- **Consecutive vacation periods** - each employee's vacation is allocated as a single continuous block
- **Two-group split** - employees are divided into two equal groups balanced by weekly_target_hours, one taking vacation in the first half of the period, the other in the second half
- **Fair shift distribution** - shifts are distributed equitably among working employees, respecting working hour constraints

## Requirements

- Python 3.6 or higher
- **openpyxl** (for Excel export): `pip install openpyxl`

## Installation

Clone this repository:
```bash
git clone https://github.com/koksbango/Sommerferieplan.git
cd Sommerferieplan
```

Install dependencies (optional, for Excel export):
```bash
pip install -r requirements.txt
```

## Usage

### Vacation Scheduler (Recommended)

Optimizes vacation allocation to maximize days while maintaining coverage:

```bash
python3 vacation_scheduler.py
```

**Custom parameters:**
```bash
python3 vacation_scheduler.py employees.csv coverage.csv 2026-06-29 5 21
```
- `employees.csv`: Employee data file
- `coverage.csv`: Coverage requirements file  
- `2026-06-29`: Start date (YYYY-MM-DD) - defaults to June 29, 2026 (week 27)
- `5`: Number of weeks - defaults to 5 weeks (June 29 - August 2, weeks 27-31)
- `21`: Target vacation days per employee

**Note:** The default vacation period is Week 27 to Week 31 (June 29 - August 2, 2026), which is the standard summer vacation period.

**Example output:**
```
Vacation capacity:
  Weekdays: up to 43 employees on vacation (need 31 working)
  Weekends: up to 48 employees on vacation (need 26 working)

Theoretical maximum:
  Total vacation-day capacity: 1555
  If distributed equally: 21 days per employee

Allocating consecutive vacation blocks with equal distribution...
  Strategy: Split employees into two equal groups
  Group 1: 37 employees (total weekly hours: 1277)
  Group 2: 37 employees (total weekly hours: 1261)

  Analyzing vacation groups...
  Group 1 (vacation in first half): 37 employees
  Group 2 (vacation in second half): 37 employees
  First vacation starts: 2026-06-29
  Last vacation starts: 2026-07-16

Results:
  Total vacation days allocated: 1258
  Average per employee: 17.0 days
  Minimum: 17 days
  Maximum: 17 days

**Note:** With 5 weeks (35 days) and two equal groups, the maximum achievable is 17 consecutive days per employee while maintaining coverage. To achieve the 21-day target, consider extending the period to 6 weeks.

Shift distribution statistics:
  Employees with shifts: 74
  Shift count range: 9 - 15 shifts
  Average shifts per working employee: 14.0
  Total hours range: 93.0 - 135.0 hours
  Average hours per working employee: 123.8

Generating Excel schedule...
✓ Excel schedule saved to: vacation_schedule.xlsx
```

The scheduler automatically generates an **Excel file** (`vacation_schedule.xlsx`) with two sheets:

**Sheet 1: Vacation Schedule**
- Visual calendar with each employee in a row and dates in columns
- "V" marks vacation days (green background)
- **Shift assignments** shown for working days (e.g., "DV", "A1", "FN")
- Working days have beige background
- Weekend dates highlighted with red headers
- Total vacation days per employee in rightmost column
- Bottom row shows daily count of employees on vacation

**Sheet 2: Shift Coverage**
- Daily shift requirements for the entire 5-week period
- Shows for each date:
  - Day of week (weekends highlighted)
  - Shift type and time range
  - Required number of employees
  - Required skill (or "Any" for general coverage)
  - Available employees (color-coded: red=insufficient, yellow=exact, green=good)
- Helps identify coverage gaps when employees are on vacation

### Vacation Calculator

Analyzes coverage requirements and shows theoretical maximums:

```bash
python3 vacation_calculator.py
```

**Custom parameters:**
```bash
python3 vacation_calculator.py employees.csv shifts.csv coverage.csv 2024-06-01 2024-08-31
```

## Input File Formats

### Employees CSV

The employees file lists each employee with their ID, name, working hours, and skills.

Format:
```csv
id,name,weekly_target_hours,max_hours_per_week,min_rest_hours_between_shifts,skills
1,"test1",37,48,11,"F;SK;AK1;T1;SIF"
2,"test2",37,48,11,"F;SK;AK1;T1;SIF"
```

- `id`: Employee ID
- `name`: Employee name
- `weekly_target_hours`: Target working hours per week
- `max_hours_per_week`: Maximum allowed hours per week
- `min_rest_hours_between_shifts`: Minimum rest hours required between shifts
- `skills`: Semicolon-separated list of skills (e.g., "F;SK;AK1;T1;SIF")

### Shifts CSV

The shifts file defines available shift types with their times and categories.

Format:
```csv
id,name,start,end,cat
1,FD,07:00,15:15,Day
2,FD2,07:00,19:00,Day
3,DV,07:15,15:30,Day
```

- `id`: Shift ID
- `name`: Shift name/code
- `start`: Shift start time (HH:MM format)
- `end`: Shift end time (HH:MM format)
- `cat`: Shift category (Day/Evening/Night)

### Coverage CSV

The coverage file defines required coverage by shift type and required skills.

Format:
```csv
type,shift_id,required,required_skills
"Weekday","FD",1,"F"
"Weekday","DV",4,"None"
"Weekday","DV",2,"SK"
"Weekend","FD",1,"F"
```

- `type`: Day type ("Weekday" or "Weekend")
- `shift_id`: Reference to shift name from shifts.csv
- `required`: Number of employees required
- `required_skills`: Required skill ("None" for general coverage, or specific skill code)

## Output

The script outputs a report showing the maximum vacation days each employee can take:

```
