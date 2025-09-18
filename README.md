# README – Total Shifts Missed (Sick and Injured)

**Report File:**  
`Sick_and_Injured.xlsx`  

---

## Overview
This report displays total hours missed due to sickness or injury for the calendar year for all uniformed members of the Rochester Fire Department.
---

## Sick or Injured – Definition
Sick or injured time is defined by these schedule type descriptions:
- Sick  
- Off Duty Injury  
- Pending On Duty Injury  
- FEMA Sick  

---

## Highlighting Rules

### Yellow Highlight
Members are highlighted **yellow** when their **`Shift_Hours`** falls within these ranges:
- **Years of service ≥ 5 years:** between **864** and **1104** hours (≈ 36–46 shifts)  
- **Years of service < 5 years:** between **552** and **864** hours (≈ 23–36 shifts)  

### Red Highlight
Members are highlighted **red** when their **`Shift_Hours`** exceeds:
- **Years of service ≥ 5 years:** **> 1104 hours** (≈ 46 shifts)  
- **Years of service < 5 years:** **> 552 hours** (≈ 23 shifts)  

---

## New Features
- **Automated Email Alerts**  
  When a member enters a red or yellow category and automated email is sent to: 
    * Case Manager
    * rfd-payroll@CityofRochester.Gov
    * Payroll Supervisor
    * CAR2, Executive Deputy Chief of Operations
    * CAR3, Executive Deputy Chief of Administration

- **Category Definitions**  
  - **Red:** Member exceeded allocated sick hours and could go into no-pay status.  
  - **Yellow:** Member is within 10 working shifts (240 hours) of moving into no-pay status.  

- **Additional Reporting Columns**  
  - Remaining hours until no-pay.  
  - "Went yellow" timestamp (date member first entered yellow category).  


---

✅ This README reflects the reviewer’s notes:  
- Conditions now based on **hours missed** (`Total_Hours_Missed`) instead of shifts.  
- Numeric thresholds converted to hours using an assumed **24-hour shift length** (e.g., 1 shift = 24 hours).  
