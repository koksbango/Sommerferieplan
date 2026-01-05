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
from typing import Dict, List, Set, Tuple
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
    
    # Initialize vacation assignments
    vacation_schedule = {emp.name: [] for emp in employees}
    
    # Track who's on vacation each day
    vacation_by_date = {date: set() for date in dates}
    
    # Improved allocation algorithm: iteratively assign vacation to employee with fewest days
    # Continue until we can't add more vacation days
    max_iterations = total_vacation_capacity + 1000  # Safety limit
    
    for iteration in range(max_iterations):
        # Find employee with fewest vacation days who can still get more
        employees_sorted = sorted(employees, key=lambda e: len(vacation_schedule[e.name]))
        
        assigned_this_iteration = False
        
        for emp in employees_sorted:
            # Stop if this employee has reached the target (or more)
            current_vacation = len(vacation_schedule[emp.name])
            if current_vacation >= target_days_per_employee:
                continue
            
            # Find days where this employee can take vacation
            for date in dates:
                # Skip if already on vacation
                if emp.name in vacation_by_date[date]:
                    continue
                
                is_weekend = date.weekday() in (5, 6)
                max_vacation_today = max_vacation_weekend if is_weekend else max_vacation_weekday
                requirements = coverage_weekend if is_weekend else coverage_weekday
                _, skill_requirements = calculate_min_employees_needed(requirements)
                
                # Check if we have capacity for more vacation
                current_vacation_count = len(vacation_by_date[date])
                if current_vacation_count >= max_vacation_today:
                    continue
                
                # Check if removing this employee still allows coverage
                employees_working = [e for e in employees 
                                    if e.name not in vacation_by_date[date] and e.name != emp.name]
                
                total_needed = weekday_total if not is_weekend else weekend_total
                
                if can_cover_with_employees(employees_working, total_needed, skill_requirements):
                    # Assign vacation
                    vacation_schedule[emp.name].append(date)
                    vacation_by_date[date].add(emp.name)
                    assigned_this_iteration = True
                    break  # Move to next employee
            
            if assigned_this_iteration:
                break  # Start new iteration from least-assigned employee
        
        # If we couldn't assign any vacation in this iteration, we're done
        if not assigned_this_iteration:
            break
        
        # Check if everyone has reached target
        min_vacation = min(len(v) for v in vacation_schedule.values())
        max_vacation = max(len(v) for v in vacation_schedule.values())
        
        # Stop if we've reached a good distribution (difference <= 1 day and min >= target)
        if min_vacation >= target_days_per_employee and (max_vacation - min_vacation) <= 1:
            break
    
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


def export_schedule_to_excel(
    vacation_schedule: Dict[str, List[datetime]],
    employees: List[Employee],
    start_date: datetime,
    num_weeks: int,
    filename: str = "vacation_schedule.xlsx"
) -> str:
    """Export vacation schedule to Excel file with visual calendar view.
    
    Args:
        vacation_schedule: Dict mapping employee name -> list of vacation dates
        employees: List of all employees
        start_date: Start date of vacation period
        num_weeks: Number of weeks in the period
        filename: Output filename (default: vacation_schedule.xlsx)
    
    Returns:
        Path to the generated file
    """
    if not OPENPYXL_AVAILABLE:
        print("\nWarning: openpyxl not available. Cannot generate Excel file.")
        print("Install with: pip install openpyxl")
        return None
    
    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Vacation Schedule"
    
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
    
    # Write header row (dates)
    ws.cell(1, 1, "Employee").fill = header_fill
    ws.cell(1, 1).font = header_font
    ws.cell(1, 1).alignment = center_align
    ws.cell(1, 1).border = border
    ws.column_dimensions['A'].width = 20
    
    for col_idx, date in enumerate(dates, start=2):
        cell = ws.cell(1, col_idx)
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
        
        ws.column_dimensions[cell.column_letter].width = 8
    
    # Add "Total" column
    total_col = len(dates) + 2
    ws.cell(1, total_col, "Total").fill = header_fill
    ws.cell(1, total_col).font = header_font
    ws.cell(1, total_col).alignment = center_align
    ws.cell(1, total_col).border = border
    ws.column_dimensions[ws.cell(1, total_col).column_letter].width = 8
    
    # Convert vacation_schedule to set of dates for faster lookup
    vacation_dates_by_employee = {
        name: set(dates) for name, dates in vacation_schedule.items()
    }
    
    # Write employee rows
    for row_idx, emp in enumerate(sorted(employees, key=lambda e: e.name), start=2):
        # Employee name
        name_cell = ws.cell(row_idx, 1, emp.name)
        name_cell.border = border
        name_cell.alignment = Alignment(vertical='center')
        
        vacation_count = 0
        
        # Mark vacation days
        for col_idx, date in enumerate(dates, start=2):
            cell = ws.cell(row_idx, col_idx)
            cell.border = border
            cell.alignment = center_align
            
            if date in vacation_dates_by_employee.get(emp.name, set()):
                cell.value = "V"
                cell.fill = vacation_fill
                cell.font = Font(bold=True)
                vacation_count += 1
            else:
                cell.value = ""
                cell.fill = working_fill
        
        # Total vacation days
        total_cell = ws.cell(row_idx, total_col, vacation_count)
        total_cell.border = border
        total_cell.alignment = center_align
        total_cell.font = Font(bold=True)
    
    # Add summary row
    summary_row = len(employees) + 3
    ws.cell(summary_row, 1, "Employees on vacation:").font = Font(bold=True)
    ws.cell(summary_row, 1).alignment = Alignment(vertical='center')
    
    for col_idx, date in enumerate(dates, start=2):
        count = sum(1 for emp in employees 
                   if date in vacation_dates_by_employee.get(emp.name, set()))
        cell = ws.cell(summary_row, col_idx, count)
        cell.alignment = center_align
        cell.font = Font(bold=True)
        cell.border = border
        
        # Color code based on count
        if count > 0:
            # Gradient from light to dark green
            intensity = min(count / 10, 1.0)
            green_val = int(144 + (0 - 144) * intensity)
            hex_color = f"90{green_val:02X}90"
            cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
    
    # Freeze panes
    ws.freeze_panes = 'B2'
    
    # Save workbook
    wb.save(filename)
    return filename



def main():
    """Main entry point."""
    employees_file = "employees.csv"
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
