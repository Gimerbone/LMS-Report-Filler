# LMS-Report-Filler
Fill a class grade Excel file with weekly Code Runner CSV results.

Usage:
```bash
    python fill_grades.py <grades.csv> <class.xlsx> <week>
```

Examples:
```bash
    python fill_grades.py week1_1.csv IF-49-01.xlsx 1.1
    python fill_grades.py fungsi.csv IF-49-01.xlsx 2
    python fill_grades.py "array.csv" IF-49-01.xlsx 4
```

The <week> argument matches the PEKAN label in row 4 of the Excel file.
Accepted forms: "1.1", "1.2", "1", "2", "3", "4", "5", "10", "11-12".