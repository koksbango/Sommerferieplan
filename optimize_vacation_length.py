#!/usr/bin/env python3
"""
Vacation Length Optimizer

This script iteratively tests different vacation lengths to find the optimal
balance between vacation days and employee workload, ensuring that max_hours_per_week
constraints are respected (avoiding Tier 3 assignments as much as possible).
"""

import sys
from datetime import datetime
from vacation_scheduler import (
    load_employees,
    load_coverage,
    load_shifts,
    get_coverage_requirements_by_type,
    optimize_vacation_schedule,
    export_schedule_to_excel
)


def analyze_workload(vacation_schedule, employees, num_weeks):
    """
    Analyze the workload distribution to check for Tier 3 usage
    and max_hours_per_week violations.
    
    Args:
        vacation_schedule: Dict mapping employee name -> list of vacation dates
        employees: List of Employee objects
        num_weeks: Number of weeks in the period
    
    Returns a dict with workload statistics.
    """
    stats = {
        'total_employees': len(employees),
        'tier3_warnings': 0,
        'max_hours_violations': 0,
        'max_weekly_hours': 0,
        'avg_weekly_hours': 0,
        'employees_over_target': 0
    }
    
    # Count working days per employee
    total_days = num_weeks * 7
    working_days = {}
    for emp in employees:
        vacation_days = len(vacation_schedule.get(emp.name, []))
        working_days[emp.name] = total_days - vacation_days
    
    # Calculate estimated weekly hours (rough estimate)
    # Assuming 8 hours per shift and working days/7 * 5 weekdays
    weekly_hours_list = []
    for emp in employees:
        if working_days[emp.name] > 0:
            # Rough estimate: working_days * 8 hours / number of weeks
            estimated_weekly = (working_days[emp.name] * 8.0) / num_weeks
            weekly_hours_list.append(estimated_weekly)
            
            if estimated_weekly > emp.max_hours_per_week:
                stats['max_hours_violations'] += 1
            if estimated_weekly > emp.weekly_target_hours:
                stats['employees_over_target'] += 1
            
            stats['max_weekly_hours'] = max(stats['max_weekly_hours'], estimated_weekly)
    
    if weekly_hours_list:
        stats['avg_weekly_hours'] = sum(weekly_hours_list) / len(weekly_hours_list)
    
    return stats


def test_vacation_length(employees, coverage_weekday, coverage_weekend, 
                        start_date, num_weeks, target_days):
    """
    Test a specific vacation length and return workload statistics.
    """
    print(f"\n{'='*70}")
    print(f"Testing {target_days} vacation days...")
    print(f"{'='*70}")
    
    # Run optimization
    vacation_schedule = optimize_vacation_schedule(
        employees,
        coverage_weekday,
        coverage_weekend,
        start_date,
        num_weeks,
        target_days
    )
    
    # Calculate actual vacation days achieved
    vacation_days_per_employee = {}
    for emp in employees:
        days = len(vacation_schedule.get(emp.name, []))
        vacation_days_per_employee[emp.name] = days
    
    min_days = min(vacation_days_per_employee.values())
    max_days = max(vacation_days_per_employee.values())
    avg_days = sum(vacation_days_per_employee.values()) / len(vacation_days_per_employee)
    
    # Analyze workload
    stats = analyze_workload(vacation_schedule, employees, num_weeks)
    
    result = {
        'target_days': target_days,
        'achieved_min': min_days,
        'achieved_max': max_days,
        'achieved_avg': avg_days,
        'vacation_schedule': vacation_schedule,
        'workload_stats': stats
    }
    
    print(f"\nVacation Days Achieved:")
    print(f"  Min: {min_days}, Max: {max_days}, Avg: {avg_days:.1f}")
    print(f"\nWorkload Statistics:")
    print(f"  Max weekly hours: {stats['max_weekly_hours']:.1f}h")
    print(f"  Avg weekly hours: {stats['avg_weekly_hours']:.1f}h")
    print(f"  Employees over target: {stats['employees_over_target']}/{stats['total_employees']}")
    print(f"  Max hours violations: {stats['max_hours_violations']}/{stats['total_employees']}")
    
    return result


def find_optimal_vacation_length(employees, coverage_weekday, coverage_weekend,
                                 start_date, num_weeks, 
                                 min_days=14, max_days=21):
    """
    Find the optimal vacation length that maximizes vacation days while
    keeping workload acceptable (no max_hours_per_week violations).
    """
    print(f"\n{'='*70}")
    print(f"OPTIMIZING VACATION LENGTH")
    print(f"{'='*70}")
    print(f"Searching range: {min_days}-{max_days} days")
    print(f"Goal: Maximize vacation days while respecting max_hours_per_week")
    
    best_result = None
    all_results = []
    
    # Test from max down to min to find the highest acceptable
    for target_days in range(max_days, min_days - 1, -1):
        result = test_vacation_length(
            employees, coverage_weekday, coverage_weekend,
            start_date, num_weeks, target_days
        )
        all_results.append(result)
        
        # Check if this is acceptable (no max_hours violations)
        if result['workload_stats']['max_hours_violations'] == 0:
            if best_result is None:
                best_result = result
                print(f"\n✓ Found acceptable solution at {target_days} days!")
                break
    
    if best_result is None:
        # No solution found without violations, take the best available
        best_result = min(all_results, 
                         key=lambda r: (r['workload_stats']['max_hours_violations'], 
                                       -r['achieved_avg']))
        print(f"\n⚠ No solution without violations found.")
        print(f"   Best compromise: {best_result['achieved_avg']:.0f} days " +
              f"with {best_result['workload_stats']['max_hours_violations']} violations")
    
    return best_result, all_results


def main():
    """Main entry point."""
    employees_file = "employees.csv"
    shifts_file = "shifts.csv"
    coverage_file = "coverage.csv"
    
    # Vacation period parameters
    start_year = 2026
    start_month = 6
    start_day = 29
    num_weeks = 5  # 5-week range (35 days)
    
    # Parse command line arguments
    if len(sys.argv) > 1:
        employees_file = sys.argv[1]
    if len(sys.argv) > 2:
        coverage_file = sys.argv[2]
    
    print(f"Loading employees from: {employees_file}")
    employees = load_employees(employees_file)
    print(f"  Loaded {len(employees)} employees")
    
    print(f"Loading shift definitions from: {shifts_file}")
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
    
    # Find optimal vacation length
    best_result, all_results = find_optimal_vacation_length(
        employees, coverage_weekday, coverage_weekend,
        start_date, num_weeks,
        min_days=14, max_days=21
    )
    
    # Print summary
    print(f"\n{'='*70}")
    print(f"OPTIMIZATION SUMMARY")
    print(f"{'='*70}")
    print(f"\nAll tested scenarios:")
    print(f"{'Target':<10} {'Achieved':<12} {'Max Hours':<12} {'Violations':<12}")
    print(f"{'-'*50}")
    for r in all_results:
        achieved = f"{r['achieved_avg']:.1f}"
        max_hours = f"{r['workload_stats']['max_weekly_hours']:.1f}h"
        violations = str(r['workload_stats']['max_hours_violations'])
        marker = " ✓" if r['workload_stats']['max_hours_violations'] == 0 else ""
        print(f"{r['target_days']:<10} {achieved:<12} {max_hours:<12} {violations:<12}{marker}")
    
    print(f"\n{'='*70}")
    print(f"RECOMMENDED SOLUTION:")
    print(f"{'='*70}")
    print(f"Vacation days: {best_result['achieved_avg']:.0f} days per employee")
    print(f"Max weekly hours: {best_result['workload_stats']['max_weekly_hours']:.1f}h")
    print(f"Avg weekly hours: {best_result['workload_stats']['avg_weekly_hours']:.1f}h")
    print(f"Employees over target: {best_result['workload_stats']['employees_over_target']}/{best_result['workload_stats']['total_employees']}")
    print(f"Max hours violations: {best_result['workload_stats']['max_hours_violations']}/{best_result['workload_stats']['total_employees']}")
    
    # Generate Excel with optimal solution
    print(f"\nGenerating Excel schedule with {best_result['achieved_avg']:.0f} vacation days...")
    excel_file = export_schedule_to_excel(
        best_result['vacation_schedule'],
        employees,
        coverage_weekday,
        coverage_weekend,
        shifts,
        start_date,
        num_weeks,
        "vacation_schedule_optimized.xlsx"
    )
    
    if excel_file:
        print(f"✓ Excel schedule saved to: {excel_file}")
    else:
        print("✗ Could not generate Excel file (openpyxl not available)")


if __name__ == "__main__":
    main()
