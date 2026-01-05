#!/usr/bin/env python3
"""
Summer Vacation Calculator for ATC Employees

This script calculates how many days of summer vacation each employee can take
based on shift requirements and employee skills.
"""

import csv
import sys
from collections import defaultdict
from datetime import datetime
from typing import Dict, List, Set, Tuple


class Employee:
    """Represents an employee with their skills."""
    
    def __init__(self, name: str, skills: Set[str]):
        self.name = name
        self.skills = skills
    
    def __repr__(self):
        return f"Employee({self.name}, {self.skills})"


class ShiftRequirement:
    """Represents a shift requirement for a specific date and skill."""
    
    def __init__(self, date: str, skill: str, min_coverage: int):
        self.date = date
        self.skill = skill
        self.min_coverage = min_coverage
    
    def __repr__(self):
        return f"ShiftRequirement({self.date}, {self.skill}, {self.min_coverage})"


def load_employees(filepath: str) -> List[Employee]:
    """Load employees from CSV file.
    
    Expected format:
    name,skills
    Alice,tower,radar
    Bob,tower
    
    Or with skills in one column:
    name,skill1,skill2,skill3
    Alice,tower,radar,
    """
    employees = []
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            headers = next(reader)  # Skip header row
            
            for row in reader:
                if not row or not row[0].strip():
                    continue
                    
                name = row[0].strip()
                # All columns after name are skills
                skills = set()
                for i in range(1, len(row)):
                    skill = row[i].strip()
                    if skill:
                        skills.add(skill)
                
                employees.append(Employee(name, skills))
    except FileNotFoundError:
        print(f"Error: File '{filepath}' not found.", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error reading employees file: {e}", file=sys.stderr)
        sys.exit(1)
    
    return employees


def load_shifts(filepath: str) -> List[ShiftRequirement]:
    """Load shift requirements from CSV file.
    
    Expected format:
    date,skill,min_coverage
    2024-06-01,tower,2
    2024-06-01,radar,1
    """
    shifts = []
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                date = row['date'].strip()
                skill = row['skill'].strip()
                min_coverage = int(row['min_coverage'].strip())
                shifts.append(ShiftRequirement(date, skill, min_coverage))
    except FileNotFoundError:
        print(f"Error: File '{filepath}' not found.", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error reading shifts file: {e}", file=sys.stderr)
        sys.exit(1)
    
    return shifts


def get_unique_dates(shifts: List[ShiftRequirement]) -> List[str]:
    """Get sorted list of unique dates from shift requirements."""
    dates = sorted(set(shift.date for shift in shifts))
    return dates


def get_skill_requirements_by_date(shifts: List[ShiftRequirement]) -> Dict[str, Dict[str, int]]:
    """Organize shift requirements by date and skill.
    
    Returns:
        Dict mapping date -> skill -> minimum coverage required
    """
    requirements = defaultdict(dict)
    for shift in shifts:
        requirements[shift.date][shift.skill] = shift.min_coverage
    return dict(requirements)


def can_cover_requirements(
    employees_available: List[Employee],
    requirements: Dict[str, int]
) -> bool:
    """Check if available employees can cover the requirements for a given day.
    
    Args:
        employees_available: List of employees working that day
        requirements: Dict mapping skill -> minimum coverage needed
    
    Returns:
        True if requirements can be met, False otherwise
    """
    # Count how many employees have each required skill
    skill_counts = defaultdict(int)
    for emp in employees_available:
        for skill in emp.skills:
            skill_counts[skill] += 1
    
    # Check if each requirement is met
    for skill, min_coverage in requirements.items():
        if skill_counts[skill] < min_coverage:
            return False
    
    return True


def calculate_max_vacation_days(
    employees: List[Employee],
    shifts: List[ShiftRequirement]
) -> Dict[str, int]:
    """Calculate maximum vacation days each employee can take.
    
    Strategy:
    - For each employee, try to maximize their vacation days
    - A day is a possible vacation day if all shift requirements can still be met
      without that employee
    
    Returns:
        Dict mapping employee name -> maximum vacation days
    """
    dates = get_unique_dates(shifts)
    requirements_by_date = get_skill_requirements_by_date(shifts)
    
    vacation_days = {}
    
    for employee in employees:
        # Count how many days this employee can be absent
        possible_vacation_days = 0
        
        for date in dates:
            if date not in requirements_by_date:
                # No requirements for this date, employee can take vacation
                possible_vacation_days += 1
                continue
            
            # Check if requirements can be met without this employee
            other_employees = [e for e in employees if e.name != employee.name]
            requirements = requirements_by_date[date]
            
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
    
    # Parse command line arguments if provided
    if len(sys.argv) > 1:
        employees_file = sys.argv[1]
    if len(sys.argv) > 2:
        shifts_file = sys.argv[2]
    
    print(f"Loading employees from: {employees_file}")
    employees = load_employees(employees_file)
    print(f"  Loaded {len(employees)} employees")
    
    print(f"Loading shift requirements from: {shifts_file}")
    shifts = load_shifts(shifts_file)
    print(f"  Loaded {len(shifts)} shift requirements")
    
    if not employees:
        print("Error: No employees loaded", file=sys.stderr)
        sys.exit(1)
    
    if not shifts:
        print("Error: No shift requirements loaded", file=sys.stderr)
        sys.exit(1)
    
    # Calculate vacation days
    dates = get_unique_dates(shifts)
    total_days = len(dates)
    
    print(f"\nCalculating vacation days for {total_days} days...")
    vacation_days = calculate_max_vacation_days(employees, shifts)
    
    # Print results
    print_results(vacation_days, total_days)


if __name__ == "__main__":
    main()
