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

Run the script with default CSV files (`employees.csv` and `shifts.csv`):

```bash
python3 vacation_calculator.py
```

### Custom CSV Files

Specify custom input files:

```bash
python3 vacation_calculator.py path/to/employees.csv path/to/shifts.csv
```

## Input File Formats

### Employees CSV

The employees file should list each employee with their skills/qualifications.

Format:
```csv
name,skills
Alice,tower,radar
Bob,tower
Carol,radar,ground
```

- First column: Employee name
- Remaining columns: Skills (one skill per column, as many columns as needed)

### Shifts CSV

The shifts file defines coverage requirements for each day and skill.

Format:
```csv
date,skill,min_coverage
2024-06-01,tower,2
2024-06-01,radar,1
2024-06-01,ground,1
2024-06-02,tower,2
```

- `date`: Date in YYYY-MM-DD format
- `skill`: Skill type required
- `min_coverage`: Minimum number of employees with this skill needed that day

## Output

The script outputs a report showing the maximum vacation days each employee can take:

```
============================================================
SUMMER VACATION ALLOCATION RESULTS
============================================================

Total days in summer period: 7

Maximum vacation days per employee:
------------------------------------------------------------
  Alice               :   5 days ( 71.4%)
  Bob                 :   7 days (100.0%)
  Carol               :   3 days ( 42.9%)
============================================================
```

## Algorithm

The calculator uses a greedy approach:

1. For each employee, iterate through all days in the vacation period
2. For each day, check if shift requirements can still be met without that employee
3. Count the days where the employee's absence doesn't violate coverage requirements
4. The count represents the maximum vacation days that employee can take

**Note**: The results show theoretical maximums. In practice, employees cannot take all their possible days simultaneously. Actual vacation scheduling requires coordination to ensure coverage.

## Example Files

The repository includes example CSV files:

- `employees.csv` - Sample employee list with 6 employees
- `employees_small.csv` - Constrained example with 3 employees
- `shifts.csv` - Sample week of shift requirements

## License

This project is open source and available under the MIT License.

