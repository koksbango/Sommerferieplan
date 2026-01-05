# Sommerferieplan

A Python script that calculates possible summer vacation days for all employees in ATC (Air Traffic Control).

## Overview

This tool helps plan summer vacation schedules by calculating the maximum number of vacation days each employee can take while ensuring all shift requirements are met. The calculator considers:

- Employee skills and qualifications
- Daily shift requirements by skill type
- Minimum coverage requirements for each skill

## Requirements

- Python 3.6 or higher
- No external dependencies (uses only standard library)

## Installation

Clone this repository:
```bash
git clone https://github.com/koksbango/Sommerferieplan.git
cd Sommerferieplan
```

## Usage

### Basic Usage

Run the script with default CSV files (`employees.csv`, `shifts.csv`, and `coverage.csv`):

```bash
python3 vacation_calculator.py
```

### Custom CSV Files

Specify custom input files:

```bash
python3 vacation_calculator.py path/to/employees.csv path/to/shifts.csv path/to/coverage.csv
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

## Algorithm

The calculator uses a greedy approach:

1. For each employee, iterate through all days in the vacation period (default: June 1 - August 31)
2. For each day, determine if it's a weekday or weekend
3. Check if shift coverage requirements can still be met without that employee
4. For each shift and skill requirement:
   - If the requirement is for a specific skill, count employees with that skill
   - If the requirement is "None" (general coverage), count all available employees
5. Count the days where the employee's absence doesn't violate coverage requirements
6. The count represents the maximum vacation days that employee can take

**Note**: The results show theoretical maximums. In practice, employees cannot take all their possible days simultaneously. Actual vacation scheduling requires coordination to ensure coverage.

## Example Files

The repository includes example CSV files with real ATC shift data:

- `employees.csv` - Sample employee list with 74 employees and their skills
- `shifts.csv` - Available shift types (Day/Evening/Night shifts)
- `coverage.csv` - Coverage requirements for weekdays and weekends
- `employees_small.csv` - Older constrained example (3 employees)
- `shifts_june.csv` - Older example format (for reference only)

## License

This project is open source and available under the MIT License.

