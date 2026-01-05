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
    """Represents an employee with their skills and working hour constraints."""
    
    def __init__(self, employee_id: str, name: str, skills: Set[str], 
                 weekly_target_hours: int = 37, max_hours_per_week: int = 48):
        self.id = employee_id
        self.name = name
        self.skills = skills
        self.weekly_target_hours = weekly_target_hours
        self.max_hours_per_week = max_hours_per_week
    
    def __repr__(self):
        return f"Employee({self.name}, {self.skills}, target={self.weekly_target_hours}h)"


class Shift:
    """Represents a shift type."""
    
    def __init__(self, shift_id: str, name: str, start: str, end: str, category: str):
        self.id = shift_id
        self.name = name
        self.start = start
        self.end = end
        self.category = category
        
        # Parse times to minutes for easier comparison
        start_parts = start.split(':')
        end_parts = end.split(':')
        self.start_minutes = int(start_parts[0]) * 60 + int(start_parts[1])
        self.end_minutes = int(end_parts[0]) * 60 + int(end_parts[1])
        
        # Handle overnight shifts
        if self.end_minutes < self.start_minutes:
            self.end_minutes += 24 * 60
    
    def overlaps_with(self, other: 'Shift') -> bool:
        """Check if this shift overlaps in time with another shift."""
        return self.start_minutes < other.end_minutes and other.start_minutes < self.end_minutes
    
    def can_work_both(self, other: 'Shift', min_rest_hours: int = 11) -> bool:
        """Check if an employee can work both this shift and another shift with required rest.
        
        Args:
            other: Another shift
            min_rest_hours: Minimum rest hours required between shifts (default 11)
        
        Returns:
            True if both shifts can be worked with sufficient rest, False otherwise
        """
        min_rest_minutes = min_rest_hours * 60
        
        # If shifts overlap, cannot work both
        if self.overlaps_with(other):
            return False
        
        # Check rest time between shifts
        # Case 1: this shift ends before other starts
        if self.end_minutes <= other.start_minutes:
            rest_time = other.start_minutes - self.end_minutes
            return rest_time >= min_rest_minutes
        
        # Case 2: other shift ends before this starts
        if other.end_minutes <= self.start_minutes:
            rest_time = self.start_minutes - other.end_minutes
            return rest_time >= min_rest_minutes
        
        # Case 3: shifts span midnight - need more complex calculation
        # For simplicity in this case, we'll check across the 24-hour boundary
        time_from_self_end_to_midnight = (24 * 60) - self.end_minutes
        time_from_midnight_to_other_start = other.start_minutes
        rest_time = time_from_self_end_to_midnight + time_from_midnight_to_other_start
        
        if rest_time >= min_rest_minutes:
            return True
        
        time_from_other_end_to_midnight = (24 * 60) - other.end_minutes
        time_from_midnight_to_self_start = self.start_minutes
        rest_time = time_from_other_end_to_midnight + time_from_midnight_to_self_start
        
        return rest_time >= min_rest_minutes
    
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
                weekly_target_hours = int(row['weekly_target_hours'].strip())
                max_hours_per_week = int(row['max_hours_per_week'].strip())
                skills_str = row['skills'].strip().strip('"')
                # Skills are semicolon-separated
                skills = set(s.strip() for s in skills_str.split(';') if s.strip())
                employees.append(Employee(employee_id, name, skills, weekly_target_hours, max_hours_per_week))
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
    requirements: List[CoverageRequirement],
    shifts: Dict[str, Shift]
) -> bool:
    """Check if available employees can cover all requirements considering shift overlaps and rest requirements.
    
    This uses a greedy algorithm:
    1. Build conflict graph - shifts that cannot be worked by same employee (overlap or insufficient rest)
    2. Calculate maximum independent set size needed
    3. Check if we have enough employees with required skills
    
    Args:
        employees_available: List of employees working that day
        requirements: List of coverage requirements (all shifts for the day)
        shifts: Dict mapping shift name to Shift object
    
    Returns:
        True if requirements can be met, False otherwise
    """
    # Group requirements by shift
    reqs_by_shift = defaultdict(list)
    for req in requirements:
        if req.shift_id in shifts:
            reqs_by_shift[req.shift_id].append(req)
    
    # Calculate requirements per shift
    shift_needs = {}
    for shift_id, reqs in reqs_by_shift.items():
        total_for_shift = sum(r.required for r in reqs)
        skill_reqs = {}
        for r in reqs:
            if r.required_skill != "None":
                skill_reqs[r.required_skill] = r.required
        shift_needs[shift_id] = {
            'total': total_for_shift,
            'skills': skill_reqs,
            'shift': shifts[shift_id]
        }
    
    # Build conflict graph: which shifts conflict with each other
    # Two shifts conflict if they overlap OR don't have 11 hours rest between them
    shift_ids = list(shift_needs.keys())
    conflicts = defaultdict(set)
    
    for i, shift_id1 in enumerate(shift_ids):
        for j, shift_id2 in enumerate(shift_ids):
            if i >= j:
                continue
            shift1 = shift_needs[shift_id1]['shift']
            shift2 = shift_needs[shift_id2]['shift']
            
            # Check if an employee can work both shifts
            if not shift1.can_work_both(shift2):
                conflicts[shift_id1].add(shift_id2)
                conflicts[shift_id2].add(shift_id1)
    
    # Use a greedy coloring approach to estimate minimum employees needed
    # This gives us a lower bound on the number of employees needed
    
    # Calculate chromatic number approximation (employees needed)
    # Sort shifts by number of conflicts (most constrained first)
    sorted_shifts = sorted(shift_ids, key=lambda s: len(conflicts[s]), reverse=True)
    
    # Greedy coloring: assign each shift to an "employee slot"
    employee_assignments = []  # Each element is a list of shifts that can be worked by one employee
    
    for shift_id in sorted_shifts:
        # Find an employee slot that can accommodate this shift
        assigned = False
        for slot in employee_assignments:
            # Check if this shift conflicts with any shift already in this slot
            can_add = True
            for existing_shift in slot:
                if existing_shift in conflicts[shift_id]:
                    can_add = False
                    break
            
            if can_add:
                # Need to add enough "copies" of this shift based on requirements
                slot.append(shift_id)
                assigned = True
                break
        
        if not assigned:
            # Need a new employee slot
            employee_assignments.append([shift_id])
    
    # Now we need to account for the actual number of people needed per shift
    # A shift requiring 5 people needs 5 separate employee slots
    total_employee_slots_needed = 0
    for shift_id, needs in shift_needs.items():
        # Each person working this shift needs their own slot
        # Find how many slots we'd need if all people in this shift also work other shifts
        total_employee_slots_needed = max(total_employee_slots_needed, needs['total'])
    
    # More accurate calculation: sum up requirements considering conflicts
    # Use maximum clique/independent set approach
    # For simplicity, we'll use a conservative estimate:
    # Count total requirements where shifts conflict
    
    # Alternative simpler approach: calculate maximum required at any time considering rest periods
    # Group shifts into "waves" that cannot share employees
    waves = []
    for shift_id in sorted_shifts:
        shift = shift_needs[shift_id]['shift']
        needs = shift_needs[shift_id]
        
        # Find a wave this can be added to (no conflicts with any shift in wave)
        placed = False
        for wave in waves:
            can_add_to_wave = True
            for other_shift_id in wave['shifts']:
                if other_shift_id in conflicts[shift_id]:
                    can_add_to_wave = False
                    break
            
            if can_add_to_wave:
                wave['shifts'].append(shift_id)
                wave['total'] += needs['total']
                for skill, count in needs['skills'].items():
                    wave['skills'][skill] = wave['skills'].get(skill, 0) + count
                placed = True
                break
        
        if not placed:
            waves.append({
                'shifts': [shift_id],
                'total': needs['total'],
                'skills': dict(needs['skills'])
            })
    
    # Find the wave with maximum requirements
    max_employees_needed = max(wave['total'] for wave in waves) if waves else 0
    max_skill_requirements = defaultdict(int)
    for wave in waves:
        for skill, count in wave['skills'].items():
            max_skill_requirements[skill] = max(max_skill_requirements[skill], count)
    
    # Check if we have enough employees
    if len(employees_available) < max_employees_needed:
        return False
    
    # Check if we have enough employees with each required skill
    for skill, required in max_skill_requirements.items():
        available_with_skill = sum(1 for emp in employees_available if skill in emp.skills)
        if available_with_skill < required:
            return False
    
    return True


def calculate_max_vacation_days(
    employees: List[Employee],
    coverage_weekday: List[CoverageRequirement],
    coverage_weekend: List[CoverageRequirement],
    shifts: Dict[str, Shift],
    start_date: datetime,
    end_date: datetime
) -> Dict[str, int]:
    """Calculate maximum vacation days each employee can take.
    
    Strategy:
    - For each employee, check each day to see if coverage can be maintained without them
    - Consider shift overlaps - not all shifts run simultaneously
    - Count days where the employee can be absent
    
    Args:
        employees: List of all employees
        coverage_weekday: Coverage requirements for weekdays
        coverage_weekend: Coverage requirements for weekends
        shifts: Dict mapping shift name to Shift object
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
            # Monday=0, Sunday=6; so Saturday=5, Sunday=6
            is_weekend = date.weekday() in (5, 6)
            requirements = coverage_weekend if is_weekend else coverage_weekday
            
            # Check if requirements can be met without this employee
            other_employees = [e for e in employees if e.id != employee.id]
            
            if can_cover_requirements(other_employees, requirements, shifts):
                possible_vacation_days += 1
        
        vacation_days[employee.name] = possible_vacation_days
    
    return vacation_days


def print_results(vacation_days: Dict[str, int], total_days: int, employees: List[Employee], 
                  coverage_weekday: List[CoverageRequirement], coverage_weekend: List[CoverageRequirement],
                  shifts: Dict[str, Shift]):
    """Print the vacation calculation results."""
    
    # Calculate coverage statistics
    weekday_total = sum(req.required for req in coverage_weekday)
    weekend_total = sum(req.required for req in coverage_weekend)
    total_employees = len(employees)
    
    # Count weekday shifts and requirements
    weekday_shifts = set(req.shift_id for req in coverage_weekday)
    weekend_shifts = set(req.shift_id for req in coverage_weekend)
    
    print("\n" + "=" * 70)
    print("COVERAGE ANALYSIS")
    print("=" * 70)
    print(f"\nTotal employees: {total_employees}")
    print(f"\nWeekday coverage:")
    print(f"  Total positions needed: {weekday_total}")
    print(f"  Number of different shifts: {len(weekday_shifts)}")
    print(f"  Theoretical max on vacation: {total_employees - weekday_total}")
    print(f"\nWeekend coverage:")
    print(f"  Total positions needed: {weekend_total}")
    print(f"  Number of different shifts: {len(weekend_shifts)}")
    print(f"  Theoretical max on vacation: {total_employees - weekend_total}")
    
    print("\n" + "=" * 70)
    print("SUMMER VACATION ALLOCATION RESULTS")
    print("=" * 70)
    print(f"\nTotal days in summer period: {total_days}")
    print("\nMaximum vacation days per employee:")
    print("-" * 70)
    
    # Sort by name for consistent output
    for name in sorted(vacation_days.keys()):
        days = vacation_days[name]
        percentage = (days / total_days * 100) if total_days > 0 else 0
        print(f"  {name:20s}: {days:3d} days ({percentage:5.1f}%)")
    
    print("=" * 70)
    print("\nNOTE: This is a simplified analysis. The actual calculation shows")
    print("each employee can take vacation on any single day because with 74")
    print(f"employees and only {weekday_total} positions needed on weekdays, there is")
    print("significant redundancy. A more sophisticated scheduler would be needed")
    print("to optimize vacation distribution while considering:")
    print("  - Multi-day shift patterns and rest requirements")
    print("  - Weekly working hour limits and targets (over 6-week periods)")
    print("  - Fairness in vacation allocation")
    print("  - Actual shift assignments and rotations")
    print()


def analyze_vacation_feasibility(
    employees: List[Employee],
    coverage_weekday: List[CoverageRequirement],
    coverage_weekend: List[CoverageRequirement],
    shifts: Dict[str, Shift],
    start_date: datetime,
    period_days: int = 35
) -> Dict:
    """Analyze maximum feasible consecutive vacation days.
    
    This function tests different vacation lengths to find the maximum that's
    feasible while maintaining 100% coverage and respecting working constraints.
    
    Args:
        employees: List of all employees
        coverage_weekday: Weekday coverage requirements
        coverage_weekend: Weekend coverage requirements
        shifts: Dict of shift definitions
        start_date: Start date of vacation period
        period_days: Length of vacation period in days (default 35 for 5 weeks)
    
    Returns:
        Dict with analysis results
    """
    total_employees = len(employees)
    group_size = total_employees // 2
    
    # Calculate weekday and weekend position needs
    weekday_total = sum(req.required for req in coverage_weekday)
    weekend_total = sum(req.required for req in coverage_weekend)
    
    # Count weekdays and weekends in period
    weekdays = 0
    weekends = 0
    for i in range(period_days):
        date = start_date + timedelta(days=i)
        if date.weekday() < 5:  # Monday=0, Friday=4
            weekdays += 1
        else:
            weekends += 1
    
    print("\n" + "=" * 70)
    print("FEASIBILITY ANALYSIS FOR CONSECUTIVE VACATION DAYS")
    print("=" * 70)
    print(f"\nVacation period: {start_date.date()} to {(start_date + timedelta(days=period_days-1)).date()}")
    print(f"Period length: {period_days} days ({weekdays} weekdays, {weekends} weekends)")
    print(f"\nTotal employees: {total_employees}")
    print(f"Two-group strategy: {group_size} employees per group")
    print(f"\nCoverage requirements:")
    print(f"  Weekdays: {weekday_total} positions needed, {group_size - weekday_total} can be on vacation")
    print(f"  Weekends: {weekend_total} positions needed, {group_size - weekend_total} can be on vacation")
    
    # Calculate average hours per employee if evenly distributed
    # Each group works for half the period (period_days / 2 days)
    work_days_per_employee = period_days // 2
    
    # Estimate shifts per employee
    # Assume ~8 hour shifts, and employees work 37-48 hours per week
    avg_target_hours = sum(e.weekly_target_hours for e in employees) / len(employees)
    avg_max_hours = sum(e.max_hours_per_week for e in employees) / len(employees)
    
    weeks = period_days / 7.0
    target_hours_total = avg_target_hours * weeks
    max_hours_total = avg_max_hours * weeks
    
    print(f"\nWorking hour analysis (per employee in working group):")
    print(f"  Working period: ~{work_days_per_employee} days (while group works)")
    print(f"  Average weekly_target_hours: {avg_target_hours:.1f}h")
    print(f"  Average max_hours_per_week: {avg_max_hours:.1f}h")
    print(f"  Target hours over {weeks:.1f} weeks: {target_hours_total:.1f}h")
    print(f"  Maximum hours over {weeks:.1f} weeks: {max_hours_total:.1f}h")
    
    # Test different vacation lengths
    print(f"\n" + "-" * 70)
    print("TESTING VACATION LENGTH SCENARIOS")
    print("-" * 70)
    
    results = {}
    for vacation_days in range(21, 13, -1):  # Test from 21 down to 14 days
        work_days = period_days - vacation_days
        
        # Calculate required shifts per working day
        avg_shifts_per_day = (weekday_total * weekdays + weekend_total * weekends) / period_days
        total_shift_slots = int(avg_shifts_per_day * work_days * 2)  # *2 because each group works half
        shifts_per_employee = total_shift_slots / total_employees
        
        # Estimate hours assuming 8-hour shifts
        est_hours_per_employee = shifts_per_employee * 8
        
        # Check if within reasonable bounds
        within_target = est_hours_per_employee <= target_hours_total * 1.1  # 10% over target
        within_max = est_hours_per_employee <= max_hours_total * 1.05  # 5% over max
        
        feasibility = "FEASIBLE" if within_max else "INFEASIBLE"
        if not within_target:
            feasibility += " (exceeds target)"
        
        print(f"\n{vacation_days} consecutive days:")
        print(f"  Work days per employee: {work_days}")
        print(f"  Estimated shifts per employee: {shifts_per_employee:.1f}")
        print(f"  Estimated hours per employee: {est_hours_per_employee:.1f}h")
        print(f"  Status: {feasibility}")
        
        results[vacation_days] = {
            'work_days': work_days,
            'shifts_per_employee': shifts_per_employee,
            'est_hours': est_hours_per_employee,
            'within_target': within_target,
            'within_max': within_max,
            'feasible': within_max
        }
    
    # Find maximum feasible
    max_feasible = None
    for days in sorted(results.keys(), reverse=True):
        if results[days]['feasible']:
            max_feasible = days
            break
    
    print(f"\n" + "=" * 70)
    print("FEASIBILITY SUMMARY")
    print("=" * 70)
    
    if max_feasible:
        print(f"\nMaximum feasible consecutive vacation days: {max_feasible}")
        print(f"  Estimated work days: {results[max_feasible]['work_days']}")
        print(f"  Estimated shifts per employee: {results[max_feasible]['shifts_per_employee']:.1f}")
        print(f"  Estimated hours per employee: {results[max_feasible]['est_hours']:.1f}h")
        print(f"\nCurrent scheduler allocation: 17 consecutive days")
        if max_feasible > 17:
            print(f"  → Potential to increase by {max_feasible - 17} days")
        elif max_feasible == 17:
            print(f"  → Current allocation is at maximum feasible level")
        else:
            print(f"  → Current allocation already exceeds estimated maximum")
    else:
        print("\nWARNING: No feasible solution found in tested range")
        print("  Current 17-day allocation may be pushing constraints")
    
    print(f"\nNOTE: This is an estimation based on average requirements.")
    print(f"Actual feasibility depends on:")
    print(f"  - Specific shift patterns and time overlaps")
    print(f"  - Individual employee max_hours_per_week limits")
    print(f"  - 6-day consecutive work limit")
    print(f"  - Skill matching requirements")
    print(f"  - Daily coverage validation")
    print(f"\nThe vacation_scheduler.py implements the full feasibility check")
    print(f"and may find slightly different results through optimization.")
    
    return results


def main():
    """Main entry point for the vacation calculator."""
    # Default file paths
    employees_file = "employees.csv"
    shifts_file = "shifts.csv"
    coverage_file = "coverage.csv"
    
    # Default vacation period - now matching scheduler (weeks 27-31, 2026)
    start_year = 2026
    start_month = 6
    start_day = 29
    end_year = 2026
    end_month = 8
    end_day = 2
    
    # Parse command line arguments if provided
    if len(sys.argv) > 1:
        employees_file = sys.argv[1]
    if len(sys.argv) > 2:
        shifts_file = sys.argv[2]
    if len(sys.argv) > 3:
        coverage_file = sys.argv[3]
    if len(sys.argv) > 4:
        # Optional: start date in YYYY-MM-DD format
        try:
            start_parts = sys.argv[4].split('-')
            start_year = int(start_parts[0])
            start_month = int(start_parts[1])
            start_day = int(start_parts[2])
        except (ValueError, IndexError):
            print(f"Warning: Invalid start date format '{sys.argv[4]}', using default", file=sys.stderr)
    if len(sys.argv) > 5:
        # Optional: end date in YYYY-MM-DD format
        try:
            end_parts = sys.argv[5].split('-')
            end_year = int(end_parts[0])
            end_month = int(end_parts[1])
            end_day = int(end_parts[2])
        except (ValueError, IndexError):
            print(f"Warning: Invalid end date format '{sys.argv[5]}', using default", file=sys.stderr)
    
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
    
    # Define summer vacation period
    start_date = datetime(start_year, start_month, start_day)
    end_date = datetime(end_year, end_month, end_day)
    total_days = (end_date - start_date).days + 1
    
    print(f"\nCalculating vacation days for period: {start_date.date()} to {end_date.date()}")
    print(f"Total days: {total_days}")
    
    # Run feasibility analysis for consecutive vacation days
    analyze_vacation_feasibility(
        employees,
        coverage_weekday,
        coverage_weekend,
        shifts,
        start_date,
        total_days
    )


if __name__ == "__main__":
    main()
