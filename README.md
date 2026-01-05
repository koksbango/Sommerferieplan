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
python3 vacation_scheduler.py employees.csv coverage.csv 2024-06-01 6 21
```
- `employees.csv`: Employee data file
- `coverage.csv`: Coverage requirements file  
- `2024-06-01`: Start date (YYYY-MM-DD)
- `6`: Number of weeks
- `21`: Target vacation days per employee

**Example output:**
```
Vacation capacity:
  Weekdays: up to 43 employees on vacation (need 31 working)
  Weekends: up to 48 employees on vacation (need 26 working)

Theoretical maximum:
  Total vacation-day capacity: 1555
  If distributed equally: 21 days per employee

Results:
  Total vacation days allocated: 1554
  Average per employee: 21.0 days
  Minimum: 21 days
  Maximum: 21 days

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
============================================================
SUMMER VACATION ALLOCATION RESULTS
============================================================

Total days in summer period: 92

Maximum vacation days per employee:
------------------------------------------------------------
  test1               :  92 days (100.0%)
  test2               :  92 days (100.0%)
  test3               :  85 days ( 92.4%)
  test4               :  78 days ( 84.8%)
============================================================
```

The vacation period is currently set to June 1 - August 31, 2024 (92 days). You can modify the `start_date` and `end_date` variables in the `main()` function to change this period.

## Key Insights

With the provided data (74 employees, 31 summer weekday positions, 26 weekend positions):
- **5 weeks with consecutive vacation and equal distribution**: All 74 employees get exactly **17 consecutive days**
  - Total: 1,258 vacation days allocated
  - Average: 17.0 days per employee
  - Spread: 0 days (perfect equality)
- **Consecutive vacation constraint**: All vacation days are allocated as a single continuous block per employee
- **Equal distribution**: Algorithm ensures all employees get the same number of days (within 1-day tolerance)
- **Weekday capacity**: Up to 43 employees can be on vacation simultaneously
- **Weekend capacity**: Up to 48 employees can be on vacation simultaneously

Note: The consecutive vacation and equal distribution constraints reduce total achievable days, but ensure fair and practical vacation schedules.

## Algorithm

The vacation scheduler uses an equal-distribution consecutive block allocation algorithm:

1. Calculate vacation capacity for each day type (weekday/weekend)
2. For each block size (starting from target), attempt to allocate that size to all employees
3. For each employee, find a consecutive block where:
   - Coverage requirements are met (total employees and skill requirements)
   - Daily vacation capacity is not exceeded
4. Try multiple employee orderings (20 attempts) to find equal distribution
5. Select the allocation that minimizes spread (max - min days), then maximizes minimum days per employee
5. Result: Each employee gets a single consecutive vacation period

## Example Files

The repository includes example CSV files with real ATC shift data:

- `employees.csv` - Sample employee list with 74 employees and their skills
- `shifts.csv` - Available shift types (Day/Evening/Night shifts)
- `coverage.csv` - Coverage requirements for weekdays and weekends
- `employees_small.csv` - Older constrained example (3 employees)
- `shifts_june.csv` - Older example format (for reference only)

## License

This project is open source and available under the MIT License.

