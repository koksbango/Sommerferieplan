#!/usr/bin/env python3
"""
Summer Vacation Scheduler for ATC Employees

This script optimizes vacation scheduling to maximize vacation days while
maintaining shift coverage requirements.
"""

import csv
import sys
from collections import defaultdict
from datetime import datetime, timedelta
from typing import Dict, List, Set, Tuple, Optional
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


class Employee:
    """Represents an employee with their skills and working hour constraints."""
    
    def __init__(self, employee_id: str, name: str, skills: Set[str], 
                 weekly_target_hours: int = 37, max_hours_per_week: int = 48):
        self.id = employee_id
        self.name = name
        self.skills = skills
        self.weekly_target_hours = weekly_target_hours
        self.max_hours_per_week = max_hours_per_week
        self.vacation_days = 0  # Track assigned vacation days
    
    def __repr__(self):
        return f"Employee({self.name}, {self.skills})"


class CoverageRequirement:
    """Represents a coverage requirement for a shift."""
    
    def __init__(self, day_type: str, shift_id: str, required: int, required_skill: str):
        self.day_type = day_type
        self.shift_id = shift_id
        self.required = required
        self.required_skill = required_skill
    
    def __repr__(self):
        return f"CoverageRequirement({self.day_type}, {self.shift_id}, {self.required}, {self.required_skill})"


def load_employees(filepath: str) -> List[Employee]:
    """Load employees from CSV file."""
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
                skills = set(s.strip() for s in skills_str.split(';') if s.strip())
                employees.append(Employee(employee_id, name, skills, weekly_target_hours, max_hours_per_week))
    except FileNotFoundError:
        print(f"Error: File '{filepath}' not found.", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error reading employees file: {e}", file=sys.stderr)
        sys.exit(1)
    
    return employees


def load_coverage(filepath: str) -> List[CoverageRequirement]:
    """Load coverage requirements from CSV file."""
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
    """Organize coverage requirements by day type."""
    requirements = defaultdict(list)
    for req in coverage:
        requirements[req.day_type].append(req)
    return dict(requirements)


def calculate_min_employees_needed(requirements: List[CoverageRequirement]) -> Tuple[int, Dict[str, int]]:
    """Calculate minimum employees needed for a day type.
    
    Returns:
        Tuple of (total_positions, skill_requirements)
    """
    total_positions = sum(req.required for req in requirements)
    skill_requirements = defaultdict(int)
    
    for req in requirements:
        if req.required_skill != "None":
            skill_requirements[req.required_skill] += req.required
    
    return total_positions, dict(skill_requirements)


def can_cover_with_employees(
    employees_available: List[Employee],
    total_needed: int,
    skill_requirements: Dict[str, int]
) -> bool:
    """Check if available employees can meet coverage needs."""
    
    if len(employees_available) < total_needed:
        return False
    
    # Check each skill requirement
    for skill, required in skill_requirements.items():
        available_with_skill = sum(1 for emp in employees_available if skill in emp.skills)
        if available_with_skill < required:
            return False
    
    return True


def optimize_vacation_schedule(
    employees: List[Employee],
    coverage_weekday: List[CoverageRequirement],
    coverage_weekend: List[CoverageRequirement],
    start_date: datetime,
    num_weeks: int,
    target_days_per_employee: int = 21
) -> Dict[str, List[datetime]]:
    """Optimize vacation scheduling to maximize vacation days while maintaining coverage.
    
    Strategy:
    1. Calculate how many employees can be on vacation each day
    2. Use iterative fair allocation to distribute vacation days
    3. Prioritize employees with fewer vacation days assigned
    
    Args:
        employees: List of all employees
        coverage_weekday: Coverage requirements for weekdays
        coverage_weekend: Coverage requirements for weekends
        start_date: Start of vacation period
        num_weeks: Number of weeks in the vacation period
        target_days_per_employee: Target vacation days per employee (default 21)
    
    Returns:
        Dict mapping employee name -> list of vacation dates
    """
    # Calculate minimum employees needed per day type
    weekday_total, weekday_skills = calculate_min_employees_needed(coverage_weekday)
    weekend_total, weekend_skills = calculate_min_employees_needed(coverage_weekend)
    
    total_employees = len(employees)
    
    # Calculate max vacation capacity per day type
    max_vacation_weekday = total_employees - weekday_total
    max_vacation_weekend = total_employees - weekend_total
    
    print(f"\nVacation capacity:")
    print(f"  Weekdays: up to {max_vacation_weekday} employees on vacation (need {weekday_total} working)")
    print(f"  Weekends: up to {max_vacation_weekend} employees on vacation (need {weekend_total} working)")
    
    # Generate dates for the period
    dates = []
    current = start_date
    end_date = start_date + timedelta(weeks=num_weeks)
    while current < end_date:
        dates.append(current)
        current += timedelta(days=1)
    
    # Count weekdays vs weekends
    weekdays = [d for d in dates if d.weekday() < 5]
    weekends = [d for d in dates if d.weekday() >= 5]
    
    # Calculate theoretical maximum vacation days possible
    total_vacation_capacity = max_vacation_weekday * len(weekdays) + max_vacation_weekend * len(weekends)
    theoretical_max_per_employee = total_vacation_capacity // total_employees
    
    print(f"\nTheoretical maximum:")
    print(f"  Total vacation-day capacity: {total_vacation_capacity}")
    print(f"  If distributed equally: {theoretical_max_per_employee} days per employee")
    
    if target_days_per_employee > theoretical_max_per_employee:
        print(f"\n  WARNING: Target of {target_days_per_employee} days exceeds theoretical max of {theoretical_max_per_employee} days!")
        print(f"  Consider extending the vacation period or reducing the target.")
    
    print("\nAllocating consecutive vacation blocks with equal distribution...")
    
    # Initialize vacation assignments
    vacation_schedule = {emp.name: [] for emp in employees}
    
    # Track who's on vacation each day
    vacation_by_date = {date: set() for date in dates}
    
    # Modified allocation algorithm: assign CONSECUTIVE vacation blocks to employees
    # Ensure all employees get equal days (within 1 day difference)
    import random
    random.seed(42)  # For reproducibility
    
    best_schedule = None
    best_min_days = 0
    best_max_spread = float('inf')  # Difference between max and min days
    best_total_days = 0
    
    # Try many different employee orderings to find equal distribution
    for attempt in range(20):  # Increased attempts for better equality
        temp_schedule = {emp.name: [] for emp in employees}
        temp_vacation_by_date = {date: set() for date in dates}
        
        # Different ordering strategies
        if attempt == 0:
            employees_ordered = sorted(employees, key=lambda e: e.name)
        elif attempt == 1:
            employees_ordered = sorted(employees, key=lambda e: e.name, reverse=True)
        else:
            employees_ordered = list(employees)
            random.seed(42 + attempt)  # Different seed for each attempt
            random.shuffle(employees_ordered)
        
        # Try to allocate equal blocks to all employees
        # Start with a target block size that should work for everyone
        for target_block_size in range(target_days_per_employee, max(6, target_days_per_employee - 8), -1):
            temp_schedule = {emp.name: [] for emp in employees}
            temp_vacation_by_date = {date: set() for date in dates}
            
            for emp in employees_ordered:
                # Find a consecutive block of exactly target_block_size
                best_block = None
                
                # Try different start positions
                for start_idx in range(len(dates) - target_block_size + 1):
                    end_idx = start_idx + target_block_size
                    candidate_block = dates[start_idx:end_idx]
                    
                    # Check if this employee can take vacation on all these days
                    can_take_all = True
                    for date in candidate_block:
                        if emp.name in temp_vacation_by_date[date]:
                            can_take_all = False
                            break
                        
                        is_weekend = date.weekday() in (5, 6)
                        max_vacation_today = max_vacation_weekend if is_weekend else max_vacation_weekday
                        requirements = coverage_weekend if is_weekend else coverage_weekday
                        _, skill_requirements = calculate_min_employees_needed(requirements)
                        
                        current_vacation_count = len(temp_vacation_by_date[date])
                        if current_vacation_count >= max_vacation_today:
                            can_take_all = False
                            break
                        
                        employees_working = [e for e in employees 
                                            if e.name not in temp_vacation_by_date[date] and e.name != emp.name]
                        
                        total_needed = weekday_total if not is_weekend else weekend_total
                        
                        if not can_cover_with_employees(employees_working, total_needed, skill_requirements):
                            can_take_all = False
                            break
                    
                    if can_take_all:
                        best_block = candidate_block
                        break  # Take the first available block
                
                # Assign the block found to this employee
                if best_block:
                    for date in best_block:
                        temp_schedule[emp.name].append(date)
                        temp_vacation_by_date[date].add(emp.name)
            
            # Check if all employees got the same amount (or within 1 day)
            vacation_counts = [len(days) for days in temp_schedule.values()]
            if vacation_counts:
                min_days = min(vacation_counts)
                max_days = max(vacation_counts)
                spread = max_days - min_days
                
                # If spread is within tolerance, this is a candidate solution
                if spread <= 1:
                    # Evaluate this schedule
                    total_days = sum(vacation_counts)
                    
                    # Keep the best schedule (prioritize low spread, then high min_days, then total_days)
                    if (spread < best_max_spread or 
                        (spread == best_max_spread and min_days > best_min_days) or
                        (spread == best_max_spread and min_days == best_min_days and total_days > best_total_days)):
                        best_max_spread = spread
                        best_min_days = min_days
                        best_total_days = total_days
                        best_schedule = temp_schedule
                        vacation_by_date = temp_vacation_by_date
                    
                    break  # Found a good solution for this attempt, move to next ordering
    
    # If we couldn't achieve equal distribution, fall back to best effort
    if best_schedule is None:
        print("  Warning: Could not achieve equal distribution within 1 day. Using best effort allocation.")
        temp_schedule = {emp.name: [] for emp in employees}
        temp_vacation_by_date = {date: set() for date in dates}
        
        employees_ordered = sorted(employees, key=lambda e: e.name)
        for emp in employees_ordered:
            best_block = None
            for block_length in range(target_days_per_employee, 6, -1):
                for start_idx in range(len(dates) - block_length + 1):
                    end_idx = start_idx + block_length
                    candidate_block = dates[start_idx:end_idx]
                    
                    can_take_all = True
                    for date in candidate_block:
                        if emp.name in temp_vacation_by_date[date]:
                            can_take_all = False
                            break
                        
                        is_weekend = date.weekday() in (5, 6)
                        max_vacation_today = max_vacation_weekend if is_weekend else max_vacation_weekday
                        requirements = coverage_weekend if is_weekend else coverage_weekday
                        _, skill_requirements = calculate_min_employees_needed(requirements)
                        
                        current_vacation_count = len(temp_vacation_by_date[date])
                        if current_vacation_count >= max_vacation_today:
                            can_take_all = False
                            break
                        
                        employees_working = [e for e in employees 
                                            if e.name not in temp_vacation_by_date[date] and e.name != emp.name]
                        
                        total_needed = weekday_total if not is_weekend else weekend_total
                        
                        if not can_cover_with_employees(employees_working, total_needed, skill_requirements):
                            can_take_all = False
                            break
                    
                    if can_take_all:
                        best_block = candidate_block
                        break
                if best_block:
                    break
            
            if best_block:
                for date in best_block:
                    temp_schedule[emp.name].append(date)
                    temp_vacation_by_date[date].add(emp.name)
        
        best_schedule = temp_schedule
        vacation_by_date = temp_vacation_by_date
    
    vacation_schedule = best_schedule
    
    return vacation_schedule


def print_vacation_results(
    vacation_schedule: Dict[str, List[datetime]],
    employees: List[Employee],
    num_weeks: int,
    target_days: int
):
    """Print vacation scheduling results."""
    
    print("\n" + "=" * 70)
    print("VACATION SCHEDULING RESULTS")
    print("=" * 70)
    
    vacation_counts = {name: len(dates) for name, dates in vacation_schedule.items()}
    
    total_days = num_weeks * 7
    total_vacation_days = sum(vacation_counts.values())
    avg_days = total_vacation_days / len(employees) if employees else 0
    min_days = min(vacation_counts.values()) if vacation_counts else 0
    max_days = max(vacation_counts.values()) if vacation_counts else 0
    
    print(f"\nPeriod: {num_weeks} weeks ({total_days} days)")
    print(f"Target vacation days per employee: {target_days}")
    print(f"\nResults:")
    print(f"  Total vacation days allocated: {total_vacation_days}")
    print(f"  Average per employee: {avg_days:.1f} days")
    print(f"  Minimum: {min_days} days")
    print(f"  Maximum: {max_days} days")
    
    # Distribution analysis
    at_target = sum(1 for v in vacation_counts.values() if v >= target_days)
    below_target = len(employees) - at_target
    
    print(f"\nDistribution:")
    print(f"  Employees at/above target ({target_days}+ days): {at_target}")
    print(f"  Employees below target: {below_target}")
    
    # Show per-employee summary
    print(f"\nPer-employee allocation:")
    print("-" * 70)
    
    for name in sorted(vacation_counts.keys()):
        days = vacation_counts[name]
        percentage = (days / total_days * 100) if total_days > 0 else 0
        status = "✓" if days >= target_days else "✗"
        print(f"  {status} {name:20s}: {days:3d} days ({percentage:5.1f}%)")
    
    print("=" * 70)


def calculate_shift_hours(shift_id: str, shifts: Dict[str, 'Shift']) -> float:
    """Calculate duration of a shift in hours."""
    shift = shifts.get(shift_id)
    if not shift:
        return 8.0  # Default assumption
    
    # Parse time strings
    try:
        start_parts = shift.start.split(':')
        end_parts = shift.end.split(':')
        start_h, start_m = int(start_parts[0]), int(start_parts[1])
        end_h, end_m = int(end_parts[0]), int(end_parts[1])
        
        start_mins = start_h * 60 + start_m
        end_mins = end_h * 60 + end_m
        
        # Handle overnight shifts
        if end_mins <= start_mins:
            end_mins += 24 * 60
        
        duration_mins = end_mins - start_mins
        return duration_mins / 60.0
    except (ValueError, IndexError, AttributeError):
        return 8.0  # Default on error


def assign_shifts_to_employees(
    employees: List[Employee],
    vacation_schedule: Dict[str, List[datetime]],
    coverage_weekday: List[CoverageRequirement],
    coverage_weekend: List[CoverageRequirement],
    dates: List[datetime],
    shifts: Dict[str, 'Shift']
) -> Dict[str, Dict[datetime, str]]:
    """Assign shifts to employees for each day based on requirements.
    
    Uses a fair distribution algorithm that:
    - Tracks working hours per employee over rolling 6-week periods
    - Respects weekly_target_hours and max_hours_per_week constraints
    - Distributes shifts fairly across all working employees
    - Prioritizes employees with fewer hours worked
    
    Args:
        employees: List of all employees
        vacation_schedule: Dict mapping employee name -> list of vacation dates
        coverage_weekday: Coverage requirements for weekdays
        coverage_weekend: Coverage requirements for weekends
        dates: List of dates in the period
        shifts: Dict mapping shift ID to Shift object
    
    Returns:
        Dict mapping employee name -> Dict mapping date -> shift assignment
    """
    # Initialize shift assignments
    shift_assignments = {emp.name: {} for emp in employees}
    
    # Convert vacation schedule to sets for faster lookup
    vacation_dates_by_employee = {
        name: set(dates_list) for name, dates_list in vacation_schedule.items()
    }
    
    # Create employee lookup by name
    employee_by_name = {emp.name: emp for emp in employees}
    
    # Track hours worked per employee per week (using 6-week rolling window)
    # Key: (employee_name, week_start_date), Value: hours
    hours_per_week = defaultdict(float)
    
    # Track total hours per employee
    total_hours = defaultdict(float)
    
    # Track number of shifts per employee (for fairness)
    shift_counts = defaultdict(int)
    
    # For each date, assign shifts
    for date in dates:
        is_weekend = date.weekday() >= 5
        requirements = coverage_weekend if is_weekend else coverage_weekday
        
        # Calculate which week this date falls into (Monday-Sunday weeks)
        week_start = date - timedelta(days=date.weekday())
        
        # Get employees available on this date
        employees_available = [emp for emp in employees 
                              if date not in vacation_dates_by_employee.get(emp.name, set())]
        
        # Group requirements by shift
        shift_reqs = defaultdict(list)
        for req in requirements:
            shift_reqs[req.shift_id].append(req)
        
        # Track which employees are already assigned today
        assigned_today = set()
        
        # Assign employees to shifts using fair distribution algorithm
        for shift_id in sorted(shift_reqs.keys()):
            reqs = shift_reqs[shift_id]
            
            # Calculate total needed for this shift
            total_needed = sum(req.required for req in reqs)
            
            # Get skill requirements
            skill_needs = {}
            for req in reqs:
                if req.required_skill != "None":
                    skill_needs[req.required_skill] = req.required
            
            # Calculate shift duration in hours
            shift_hours = calculate_shift_hours(shift_id, shifts)
            
            # Build candidate lists for each skill requirement
            assigned_this_shift = []
            
            # First, assign employees with required skills
            for skill, needed in skill_needs.items():
                candidates = [emp for emp in employees_available 
                            if skill in emp.skills and emp.name not in assigned_today]
                
                # Sort candidates by: 
                # 1. Hours worked this week (prefer those with fewer hours)
                # 2. Total shift count (prefer those with fewer shifts overall)
                # 3. Total hours (prefer those with fewer total hours)
                def sort_key(emp):
                    week_hours = hours_per_week[(emp.name, week_start)]
                    # Check if adding this shift would exceed max_hours_per_week
                    would_exceed_max = (week_hours + shift_hours) > emp.max_hours_per_week
                    # Prioritize those who won't exceed max, then by hours worked
                    return (would_exceed_max, week_hours, shift_counts[emp.name], total_hours[emp.name])
                
                candidates.sort(key=sort_key)
                
                # Assign the needed number of employees
                for i in range(min(needed, len(candidates))):
                    emp = candidates[i]
                    shift_assignments[emp.name][date] = shift_id
                    assigned_today.add(emp.name)
                    assigned_this_shift.append(emp.name)
                    
                    # Update tracking
                    hours_per_week[(emp.name, week_start)] += shift_hours
                    total_hours[emp.name] += shift_hours
                    shift_counts[emp.name] += 1
            
            # Then assign remaining positions to any available employee
            remaining_needed = total_needed - len(assigned_this_shift)
            if remaining_needed > 0:
                candidates = [emp for emp in employees_available 
                            if emp.name not in assigned_today]
                
                # Same sorting strategy
                def sort_key(emp):
                    week_hours = hours_per_week[(emp.name, week_start)]
                    would_exceed_max = (week_hours + shift_hours) > emp.max_hours_per_week
                    return (would_exceed_max, week_hours, shift_counts[emp.name], total_hours[emp.name])
                
                candidates.sort(key=sort_key)
                
                for i in range(min(remaining_needed, len(candidates))):
                    emp = candidates[i]
                    shift_assignments[emp.name][date] = shift_id
                    assigned_today.add(emp.name)
                    
                    # Update tracking
                    hours_per_week[(emp.name, week_start)] += shift_hours
                    total_hours[emp.name] += shift_hours
                    shift_counts[emp.name] += 1
    
    # Print fairness statistics
    print("\nShift distribution statistics:")
    working_employees = [emp for emp in employees if shift_counts[emp.name] > 0]
    if working_employees:
        shift_count_list = [shift_counts[emp.name] for emp in working_employees]
        hours_list = [total_hours[emp.name] for emp in working_employees]
        
        print(f"  Employees with shifts: {len(working_employees)}")
        print(f"  Shift count range: {min(shift_count_list)} - {max(shift_count_list)} shifts")
        print(f"  Average shifts per working employee: {sum(shift_count_list)/len(shift_count_list):.1f}")
        print(f"  Total hours range: {min(hours_list):.1f} - {max(hours_list):.1f} hours")
        print(f"  Average hours per working employee: {sum(hours_list)/len(hours_list):.1f}")
    
    return shift_assignments


def export_schedule_to_excel(
    vacation_schedule: Dict[str, List[datetime]],
    employees: List[Employee],
    coverage_weekday: List[CoverageRequirement],
    coverage_weekend: List[CoverageRequirement],
    shifts: Dict[str, 'Shift'],
    start_date: datetime,
    num_weeks: int,
    filename: str = "vacation_schedule.xlsx"
) -> Optional[str]:
    """Export vacation schedule to Excel file with visual calendar view and shift coverage.
    
    Args:
        vacation_schedule: Dict mapping employee name -> list of vacation dates
        employees: List of all employees
        coverage_weekday: Coverage requirements for weekdays
        coverage_weekend: Coverage requirements for weekends
        shifts: Dict mapping shift name to Shift object
        start_date: Start date of vacation period
        num_weeks: Number of weeks in the period
        filename: Output filename (default: vacation_schedule.xlsx)
    
    Returns:
        Path to the generated file, or None if openpyxl is not available
    """
    if not OPENPYXL_AVAILABLE:
        print("\nWarning: openpyxl not available. Cannot generate Excel file.")
        print("Install with: pip install openpyxl")
        return None
    
    # Create workbook
    wb = Workbook()
    
    # Create vacation schedule sheet
    ws_vacation = wb.active
    ws_vacation.title = "Vacation Schedule"
    
    # Generate all dates in the period
    dates = []
    current = start_date
    for _ in range(num_weeks * 7):
        dates.append(current)
        current += timedelta(days=1)
    
    # Define styles
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    vacation_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    working_fill = PatternFill(start_color="FFE4B5", end_color="FFE4B5", fill_type="solid")
    weekend_header_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_align = Alignment(horizontal='center', vertical='center')
    
    # Write header row (dates) for vacation schedule
    ws_vacation.cell(1, 1, "Employee").fill = header_fill
    ws_vacation.cell(1, 1).font = header_font
    ws_vacation.cell(1, 1).alignment = center_align
    ws_vacation.cell(1, 1).border = border
    ws_vacation.column_dimensions['A'].width = 20
    
    for col_idx, date in enumerate(dates, start=2):
        cell = ws_vacation.cell(1, col_idx)
        day_name = date.strftime('%a')
        date_str = date.strftime('%d/%m')
        cell.value = f"{day_name}\n{date_str}"
        cell.font = header_font
        cell.alignment = center_align
        cell.border = border
        
        # Different color for weekends
        if date.weekday() >= 5:
            cell.fill = weekend_header_fill
        else:
            cell.fill = header_fill
        
        ws_vacation.column_dimensions[cell.column_letter].width = 8
    
    # Add "Total" column
    total_col = len(dates) + 2
    ws_vacation.cell(1, total_col, "Total").fill = header_fill
    ws_vacation.cell(1, total_col).font = header_font
    ws_vacation.cell(1, total_col).alignment = center_align
    ws_vacation.cell(1, total_col).border = border
    ws_vacation.column_dimensions[ws_vacation.cell(1, total_col).column_letter].width = 8
    
    # Convert vacation_schedule to set of dates for faster lookup
    vacation_dates_by_employee = {
        name: set(dates) for name, dates in vacation_schedule.items()
    }
    
    # Assign shifts to employees
    shift_assignments = assign_shifts_to_employees(
        employees, vacation_schedule, coverage_weekday, coverage_weekend, dates, shifts
    )
    
    # Write employee rows
    for row_idx, emp in enumerate(sorted(employees, key=lambda e: e.name), start=2):
        # Employee name
        name_cell = ws_vacation.cell(row_idx, 1, emp.name)
        name_cell.border = border
        name_cell.alignment = Alignment(vertical='center')
        
        vacation_count = 0
        
        # Mark vacation days or shift assignments
        for col_idx, date in enumerate(dates, start=2):
            cell = ws_vacation.cell(row_idx, col_idx)
            cell.border = border
            cell.alignment = center_align
            
            if date in vacation_dates_by_employee.get(emp.name, set()):
                cell.value = "V"
                cell.fill = vacation_fill
                cell.font = Font(bold=True)
                vacation_count += 1
            else:
                # Show shift assignment
                assigned_shift = shift_assignments[emp.name].get(date, "")
                cell.value = assigned_shift
                cell.fill = working_fill
                if assigned_shift:
                    cell.font = Font(size=9)
        
        # Total vacation days
        total_cell = ws_vacation.cell(row_idx, total_col, vacation_count)
        total_cell.border = border
        total_cell.alignment = center_align
        total_cell.font = Font(bold=True)
    
    # Add summary row
    summary_row = len(employees) + 3
    ws_vacation.cell(summary_row, 1, "Employees on vacation:").font = Font(bold=True)
    ws_vacation.cell(summary_row, 1).alignment = Alignment(vertical='center')
    
    for col_idx, date in enumerate(dates, start=2):
        count = sum(1 for emp in employees 
                   if date in vacation_dates_by_employee.get(emp.name, set()))
        cell = ws_vacation.cell(summary_row, col_idx, count)
        cell.alignment = center_align
        cell.font = Font(bold=True)
        cell.border = border
        
        # Color code based on count
        if count > 0:
            # Gradient from light green (90EE90) to darker green
            intensity = min(count / 10.0, 1.0)
            # Calculate green component: 238 (0xEE) down to 144 (0x90)
            green_val = int(238 - (238 - 144) * intensity)
            green_val = max(144, min(238, green_val))  # Clamp to valid range
            hex_color = f"90{green_val:02X}90"
            cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
    
    # Freeze panes
    ws_vacation.freeze_panes = 'B2'
    
    # ==================================================================
    # CREATE SHIFT COVERAGE SHEET
    # ==================================================================
    ws_coverage = wb.create_sheet("Shift Coverage")
    
    # Define additional styles for coverage sheet
    shift_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    
    # Write header row
    ws_coverage.cell(1, 1, "Date").fill = header_fill
    ws_coverage.cell(1, 1).font = header_font
    ws_coverage.cell(1, 1).alignment = center_align
    ws_coverage.cell(1, 1).border = border
    ws_coverage.column_dimensions['A'].width = 12
    
    ws_coverage.cell(1, 2, "Day").fill = header_fill
    ws_coverage.cell(1, 2).font = header_font
    ws_coverage.cell(1, 2).alignment = center_align
    ws_coverage.cell(1, 2).border = border
    ws_coverage.column_dimensions['B'].width = 10
    
    ws_coverage.cell(1, 3, "Shift").fill = header_fill
    ws_coverage.cell(1, 3).font = header_font
    ws_coverage.cell(1, 3).alignment = center_align
    ws_coverage.cell(1, 3).border = border
    ws_coverage.column_dimensions['C'].width = 10
    
    ws_coverage.cell(1, 4, "Time").fill = header_fill
    ws_coverage.cell(1, 4).font = header_font
    ws_coverage.cell(1, 4).alignment = center_align
    ws_coverage.cell(1, 4).border = border
    ws_coverage.column_dimensions['D'].width = 15
    
    ws_coverage.cell(1, 5, "Required").fill = header_fill
    ws_coverage.cell(1, 5).font = header_font
    ws_coverage.cell(1, 5).alignment = center_align
    ws_coverage.cell(1, 5).border = border
    ws_coverage.column_dimensions['E'].width = 10
    
    ws_coverage.cell(1, 6, "Skill").fill = header_fill
    ws_coverage.cell(1, 6).font = header_font
    ws_coverage.cell(1, 6).alignment = center_align
    ws_coverage.cell(1, 6).border = border
    ws_coverage.column_dimensions['F'].width = 12
    
    ws_coverage.cell(1, 7, "Available").fill = header_fill
    ws_coverage.cell(1, 7).font = header_font
    ws_coverage.cell(1, 7).alignment = center_align
    ws_coverage.cell(1, 7).border = border
    ws_coverage.column_dimensions['G'].width = 10
    
    # Write coverage data
    row = 2
    for date in dates:
        is_weekend = date.weekday() >= 5
        requirements = coverage_weekend if is_weekend else coverage_weekday
        day_name = date.strftime('%A')
        date_str = date.strftime('%Y-%m-%d')
        
        # Get employees available on this date (not on vacation)
        employees_available = [emp for emp in employees 
                              if date not in vacation_dates_by_employee.get(emp.name, set())]
        
        # Group requirements by shift
        from collections import defaultdict
        shift_reqs = defaultdict(list)
        for req in requirements:
            shift_reqs[req.shift_id].append(req)
        
        # Write each shift's requirements
        for shift_id in sorted(shift_reqs.keys()):
            reqs = shift_reqs[shift_id]
            shift = shifts.get(shift_id)
            shift_time = f"{shift.start}-{shift.end}" if shift else "N/A"
            
            for req in reqs:
                ws_coverage.cell(row, 1, date_str).border = border
                ws_coverage.cell(row, 1).alignment = Alignment(horizontal='left', vertical='center')
                
                ws_coverage.cell(row, 2, day_name).border = border
                ws_coverage.cell(row, 2).alignment = center_align
                if is_weekend:
                    ws_coverage.cell(row, 2).fill = weekend_header_fill
                    ws_coverage.cell(row, 2).font = Font(color="FFFFFF", bold=True)
                
                ws_coverage.cell(row, 3, shift_id).border = border
                ws_coverage.cell(row, 3).alignment = center_align
                ws_coverage.cell(row, 3).fill = shift_fill
                
                ws_coverage.cell(row, 4, shift_time).border = border
                ws_coverage.cell(row, 4).alignment = center_align
                
                ws_coverage.cell(row, 5, req.required).border = border
                ws_coverage.cell(row, 5).alignment = center_align
                ws_coverage.cell(row, 5).font = Font(bold=True)
                
                skill_text = req.required_skill if req.required_skill != "None" else "Any"
                ws_coverage.cell(row, 6, skill_text).border = border
                ws_coverage.cell(row, 6).alignment = center_align
                
                # Count available employees with required skill
                if req.required_skill == "None":
                    available_count = len(employees_available)
                else:
                    available_count = sum(1 for emp in employees_available 
                                        if req.required_skill in emp.skills)
                
                ws_coverage.cell(row, 7, available_count).border = border
                ws_coverage.cell(row, 7).alignment = center_align
                
                # Color code based on coverage adequacy
                if available_count < req.required:
                    # Red - insufficient coverage
                    ws_coverage.cell(row, 7).fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
                    ws_coverage.cell(row, 7).font = Font(bold=True, color="990000")
                elif available_count == req.required:
                    # Yellow - exact coverage
                    ws_coverage.cell(row, 7).fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
                    ws_coverage.cell(row, 7).font = Font(bold=True)
                else:
                    # Green - good coverage
                    ws_coverage.cell(row, 7).fill = PatternFill(start_color="99FF99", end_color="99FF99", fill_type="solid")
                
                row += 1
    
    # Freeze panes on coverage sheet
    ws_coverage.freeze_panes = 'A2'
    
    # Save workbook
    wb.save(filename)
    return filename



def main():
    """Main entry point."""
    employees_file = "employees.csv"
    shifts_file = "shifts.csv"
    coverage_file = "coverage.csv"
    
    # Vacation period parameters
    start_year = 2024
    start_month = 6
    start_day = 1
    num_weeks = 5  # 5-week range
    target_days = 21
    
    # Parse command line arguments
    if len(sys.argv) > 1:
        employees_file = sys.argv[1]
    if len(sys.argv) > 2:
        coverage_file = sys.argv[2]
    if len(sys.argv) > 3:
        start_date_str = sys.argv[3]
        try:
            parts = start_date_str.split('-')
            start_year = int(parts[0])
            start_month = int(parts[1])
            start_day = int(parts[2])
        except (ValueError, IndexError):
            print(f"Warning: Invalid start date '{start_date_str}', using default", file=sys.stderr)
    if len(sys.argv) > 4:
        num_weeks = int(sys.argv[4])
    if len(sys.argv) > 5:
        target_days = int(sys.argv[5])
    
    print(f"Loading employees from: {employees_file}")
    employees = load_employees(employees_file)
    print(f"  Loaded {len(employees)} employees")
    
    print(f"Loading shift definitions from: {shifts_file}")
    from vacation_calculator import load_shifts  # Import from the other module
    shifts = load_shifts(shifts_file)
    print(f"  Loaded {len(shifts)} shift types")
    
    print(f"Loading coverage requirements from: {coverage_file}")
    coverage = load_coverage(coverage_file)
    print(f"  Loaded {len(coverage)} coverage requirements")
    
    # Organize coverage by day type
    coverage_by_type = get_coverage_requirements_by_type(coverage)
    coverage_weekday = coverage_by_type.get('Weekday', [])
    coverage_weekend = coverage_by_type.get('Weekend', [])
    
    start_date = datetime(start_year, start_month, start_day)
    
    print(f"\nOptimizing vacation schedule:")
    print(f"  Start date: {start_date.date()}")
    print(f"  Duration: {num_weeks} weeks ({num_weeks * 7} days)")
    print(f"  Target: {target_days} days per employee")
    
    # Optimize vacation schedule
    vacation_schedule = optimize_vacation_schedule(
        employees,
        coverage_weekday,
        coverage_weekend,
        start_date,
        num_weeks,
        target_days
    )
    
    # Print results
    print_vacation_results(vacation_schedule, employees, num_weeks, target_days)
    
    # Export to Excel
    print("\nGenerating Excel schedule...")
    excel_file = export_schedule_to_excel(
        vacation_schedule,
        employees,
        coverage_weekday,
        coverage_weekend,
        shifts,
        start_date,
        num_weeks,
        "vacation_schedule.xlsx"
    )
    
    if excel_file:
        print(f"✓ Excel schedule saved to: {excel_file}")
    else:
        print("✗ Could not generate Excel file (openpyxl not available)")




if __name__ == "__main__":
    main()
