#!/usr/bin/env python3
"""
Summer Vacation Calculator for ATC Employees

This script calculates how many days of summer vacation each employee can take
based on shift requirements and employee skills.
"""

import csv
import sys
from collections import defaultdict
from datetime import datetime, timedelta
from typing import Dict, List, Set, Tuple


class Employee:
    """Represents an employee with their skills."""
    
    def __init__(self, employee_id: str, name: str, skills: Set[str]):
        self.id = employee_id
        self.name = name
        self.skills = skills
    
    def __repr__(self):
        return f"Employee({self.name}, {self.skills})"


class Shift:
    """Represents a shift type."""
    
    def __init__(self, shift_id: str, name: str, start: str, end: str, category: str):
        self.id = shift_id
        self.name = name
        self.start = start
        self.end = end
        self.category = category
    
    def __repr__(self):
        return f"Shift({self.name}, {self.start}-{self.end}, {self.category})"


class CoverageRequirement:
    """Represents a coverage requirement for a shift."""
    
    def __init__(self, day_type: str, shift_id: str, required: int, required_skill: str):
        self.day_type = day_type  # "Weekday" or "Weekend"
        self.shift_id = shift_id
        self.required = required
        self.required_skill = required_skill  # Can be "None" for general coverage
    
    def __repr__(self):
        return f"CoverageRequirement({self.day_type}, {self.shift_id}, {self.required}, {self.required_skill})"


def load_employees(filepath: str) -> List[Employee]:
    """Load employees from CSV file.
    
    Expected format:
    id,name,weekly_target_hours,max_hours_per_week,min_rest_hours_between_shifts,skills
    1,"test1",37,48,11,"F;SK;AK1;T1;SIF"
    """
    employees = []
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                employee_id = row['id'].strip()
                name = row['name'].strip().strip('"')
                skills_str = row['skills'].strip().strip('"')
                # Skills are semicolon-separated
                skills = set(s.strip() for s in skills_str.split(';') if s.strip())
                employees.append(Employee(employee_id, name, skills))
    except FileNotFoundError:
        print(f"Error: File '{filepath}' not found.", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error reading employees file: {e}", file=sys.stderr)
        sys.exit(1)
    
    return employees


def load_shifts(filepath: str) -> Dict[str, Shift]:
    """Load shift definitions from CSV file.
    
    Expected format:
    id,name,start,end,cat
    1,FD,07:00,15:15,Day
    
    Returns:
        Dict mapping shift name -> Shift object
    """
    shifts = {}
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                shift_id = row['id'].strip()
                name = row['name'].strip()
                start = row['start'].strip()
                end = row['end'].strip()
                category = row['cat'].strip()
                shifts[name] = Shift(shift_id, name, start, end, category)
    except FileNotFoundError:
        print(f"Error: File '{filepath}' not found.", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error reading shifts file: {e}", file=sys.stderr)
        sys.exit(1)
    
    return shifts


def load_coverage(filepath: str) -> List[CoverageRequirement]:
    """Load coverage requirements from CSV file.
    
    Expected format:
    type,shift_id,required,required_skills
    "Weekday","FD",1,"F"
    "Weekday","DV",4,"None"
    """
    coverage = []
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                day_type = row['type'].strip().strip('"')
                shift_id = row['shift_id'].strip().strip('"')
                required = int(row['required'].strip())
                required_skill = row['required_skills'].strip().strip('"')
                coverage.append(CoverageRequirement(day_type, shift_id, required, required_skill))
    except FileNotFoundError:
        print(f"Error: File '{filepath}' not found.", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error reading coverage file: {e}", file=sys.stderr)
        sys.exit(1)
    
    return coverage


def get_coverage_requirements_by_type(coverage: List[CoverageRequirement]) -> Dict[str, List[CoverageRequirement]]:
    """Organize coverage requirements by day type (Weekday/Weekend).
    
    Returns:
        Dict mapping day_type -> list of requirements
    """
    requirements = defaultdict(list)
    for req in coverage:
        requirements[req.day_type].append(req)
    return dict(requirements)


def can_cover_requirements(
    employees_available: List[Employee],
    requirements: List[CoverageRequirement]
) -> bool:
    """Check if available employees can cover all requirements.
    
    Args:
        employees_available: List of employees working that day
        requirements: List of coverage requirements
    
    Returns:
        True if requirements can be met, False otherwise
    """
    # Group requirements by shift
    shift_requirements = defaultdict(list)
    for req in requirements:
        shift_requirements[req.shift_id].append(req)
    
    # For each shift, check if we have enough coverage
    for shift_id, reqs in shift_requirements.items():
        # Check each specific skill requirement
        for req in reqs:
            if req.required_skill == "None":
                # General coverage - count all available employees
                available_count = len(employees_available)
            else:
                # Specific skill required - count employees with that skill
                available_count = sum(1 for emp in employees_available if req.required_skill in emp.skills)
            
            if available_count < req.required:
                return False
    
    return True


def calculate_max_vacation_days(
    employees: List[Employee],
    coverage_weekday: List[CoverageRequirement],
    coverage_weekend: List[CoverageRequirement],
    start_date: datetime,
    end_date: datetime
) -> Dict[str, int]:
    """Calculate maximum vacation days each employee can take.
    
    Strategy:
    - For each employee, try to maximize their vacation days
    - A day is a possible vacation day if all shift requirements can still be met
      without that employee
    - Consider weekday vs weekend requirements
    
    Args:
        employees: List of all employees
        coverage_weekday: Coverage requirements for weekdays
        coverage_weekend: Coverage requirements for weekends
        start_date: Start of vacation period
        end_date: End of vacation period
    
    Returns:
        Dict mapping employee name -> maximum vacation days
    """
    vacation_days = {}
    
    # Generate list of dates in the period
    dates = []
    current = start_date
    while current <= end_date:
        dates.append(current)
        current += timedelta(days=1)
    
    for employee in employees:
        # Count how many days this employee can be absent
        possible_vacation_days = 0
        
        for date in dates:
            # Determine if it's a weekday or weekend
            # Monday=0, Sunday=6
            is_weekend = date.weekday() >= 5  # Saturday or Sunday
            requirements = coverage_weekend if is_weekend else coverage_weekday
            
            # Check if requirements can be met without this employee
            other_employees = [e for e in employees if e.id != employee.id]
            
            if can_cover_requirements(other_employees, requirements):
                possible_vacation_days += 1
        
        vacation_days[employee.name] = possible_vacation_days
    
    return vacation_days


def print_results(vacation_days: Dict[str, int], total_days: int):
    """Print the vacation calculation results."""
    print("\n" + "=" * 60)
    print("SUMMER VACATION ALLOCATION RESULTS")
    print("=" * 60)
    print(f"\nTotal days in summer period: {total_days}")
    print("\nMaximum vacation days per employee:")
    print("-" * 60)
    
    # Sort by name for consistent output
    for name in sorted(vacation_days.keys()):
        days = vacation_days[name]
        percentage = (days / total_days * 100) if total_days > 0 else 0
        print(f"  {name:20s}: {days:3d} days ({percentage:5.1f}%)")
    
    print("=" * 60)
    print("\nNote: These are maximum possible days. Actual allocation may vary")
    print("based on coordination and fairness policies.")
    print()


def main():
    """Main entry point for the vacation calculator."""
    # Default file paths
    employees_file = "employees.csv"
    shifts_file = "shifts.csv"
    coverage_file = "coverage.csv"
    
    # Parse command line arguments if provided
    if len(sys.argv) > 1:
        employees_file = sys.argv[1]
    if len(sys.argv) > 2:
        shifts_file = sys.argv[2]
    if len(sys.argv) > 3:
        coverage_file = sys.argv[3]
    
    print(f"Loading employees from: {employees_file}")
    employees = load_employees(employees_file)
    print(f"  Loaded {len(employees)} employees")
    
    print(f"Loading shift definitions from: {shifts_file}")
    shifts = load_shifts(shifts_file)
    print(f"  Loaded {len(shifts)} shift types")
    
    print(f"Loading coverage requirements from: {coverage_file}")
    coverage = load_coverage(coverage_file)
    print(f"  Loaded {len(coverage)} coverage requirements")
    
    if not employees:
        print("Error: No employees loaded", file=sys.stderr)
        sys.exit(1)
    
    if not coverage:
        print("Error: No coverage requirements loaded", file=sys.stderr)
        sys.exit(1)
    
    # Organize coverage by day type
    coverage_by_type = get_coverage_requirements_by_type(coverage)
    coverage_weekday = coverage_by_type.get('Weekday', [])
    coverage_weekend = coverage_by_type.get('Weekend', [])
    
    # Define summer vacation period (example: June 1 - August 31, 2024)
    # User can modify these dates as needed
    start_date = datetime(2024, 6, 1)
    end_date = datetime(2024, 8, 31)
    total_days = (end_date - start_date).days + 1
    
    print(f"\nCalculating vacation days for period: {start_date.date()} to {end_date.date()}")
    print(f"Total days: {total_days}")
    
    # Calculate vacation days
    vacation_days = calculate_max_vacation_days(
        employees, 
        coverage_weekday, 
        coverage_weekend,
        start_date,
        end_date
    )
    
    # Print results
    print_results(vacation_days, total_days)


if __name__ == "__main__":
    main()
