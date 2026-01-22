# Sommerferieplan

A Python script that calculates and optimizes summer vacation schedules for ATC employees.

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

Allocates vacation based on employee wishes while maintaining shift coverage:

```bash
python3 vacation_scheduler.py
```

**Requirements:**
- `vacation_wishes.csv` file is **required** - must contain vacation preferences for all employees
- See [Vacation Wishes CSV](#vacation-wishes-csv-optional) section below for format details

**Custom parameters:**
```bash
python3 vacation_scheduler.py employees.csv coverage.csv 2026-04-27 23 21
```
- `employees.csv`: Employee data file
- `coverage.csv`: Coverage requirements file  
- `2026-04-27`: Start date (YYYY-MM-DD) - defaults to April 27, 2026 (week 18)
- `23`: Number of weeks - defaults to 23 weeks (April 27 - October 4, weeks 18-40)
- `21`: Target vacation days per employee

**Note:** The vacation period spans Week 18 to Week 40 (April 27 - October 4, 2026), covering May through September. The script reads employee vacation wishes from `vacation_wishes.csv` and allocates 3 weeks per employee from their 4 prioritized week requests.

**Example output:**
```
Vacation capacity:
  Weekdays: up to 43 employees on vacation (need 31 working)
  Weekends: up to 48 employees on vacation (need 26 working)

Using vacation wish-based allocation...

Allocating vacation from wishes (target: 3 weeks per employee):
  Employees with wishes: 74
  Weekday capacity: up to 43 employees on vacation
  Weekend capacity: up to 48 employees on vacation

Vacation allocation results:
  Employees with wishes who got vacation: 74
  Average weeks allocated: 3.0
  Min weeks allocated: 3
  Max weeks allocated: 3
  Distribution:
    3 weeks: 74 employees

Results:
  Total vacation days allocated: 1554
  Average per employee: 21.0 days
  Minimum: 21 days
  Maximum: 21 days

Distribution:
  Employees at/above target (21+ days): 74
  Employees below target: 0

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

### Vacation Wishes CSV (Required)

The vacation wishes file is **required** for the vacation scheduler to run. It allows employees to prioritize their preferred vacation weeks, and the scheduler allocates 3 weeks out of each employee's 4 requested weeks while maintaining shift coverage.

Format:
```csv
employee,priority1,priority2,priority3,priority4
"test1",27,28,26,29
"test2",27,30,28,31
```

- `employee`: Employee name (must match name in employees.csv)
- `priority1`: Week number for highest priority (1 = most desired)
- `priority2`: Week number for second priority
- `priority3`: Week number for third priority
- `priority4`: Week number for lowest priority (4 = least desired)

**Week Numbers:**
- Week numbers range from 1 to 52 (ISO week numbers)
- Summer vacation period: weeks 18-40 (beginning of May to end of September)
- Requests outside the 18-40 range will be rejected with a warning

**Allocation Strategy:**
- The system allocates 3 weeks per employee from their 4 requested weeks
- Requests are processed in priority order (priority 1 first, then 2, 3, 4)
- When multiple employees request the same week, priority levels are used to resolve conflicts
- Coverage requirements and skill constraints are maintained
- All employees must have vacation wishes defined in the CSV file

## Output

The script outputs a report showing vacation allocation results:

```
