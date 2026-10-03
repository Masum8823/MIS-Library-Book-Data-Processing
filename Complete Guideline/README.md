# MIS Project — University Library Book Data Processing

A complete step-by-step **Management Information System (MIS)** project for processing university library book data using **Microsoft Excel**.

This project demonstrates how raw library data can be collected, entered, validated, stored, processed, analyzed, reported, and visualized.

---

# Table of Contents

* [1. Project Overview](#1-project-overview)
* [2. Complete Project Workflow](#2-complete-project-workflow)
* [3. Raw Data](#3-raw-data)
* [4. Data Fields](#4-data-fields)
* [5. Stage 1 — Data Collection](#5-stage-1--data-collection)
* [6. Stage 2 — Data Entry](#6-stage-2--data-entry)
* [7. Stage 3 — Data Validation](#7-stage-3--data-validation)
* [8. Stage 4 — Data Storage](#8-stage-4--data-storage)
* [9. Stage 5 — Data Processing](#9-stage-5--data-processing)
* [10. Stage 6 — Data Analysis](#10-stage-6--data-analysis)
* [11. Stage 7 — Reporting](#11-stage-7--reporting)
* [12. Stage 8 — Data Visualization](#12-stage-8--data-visualization)
* [13. Important Data Validation Finding](#13-important-data-validation-finding)
* [14. Final Analysis Summary](#14-final-analysis-summary)
* [15. Final Excel Structure](#15-final-excel-structure)
* [16. Important Excel Functions](#16-important-excel-functions)
* [17. Final Checklist](#17-final-checklist)

---

# 1. Project Overview

## Project Title

**University Library Book Data Processing**

## Project Scenario

A university wants to develop a simple **Library Management Information System**.

Raw library book information was collected from different sources. The data is not properly organized and may contain inconsistencies.

The objective of this project is to process the raw library data using Microsoft Excel and generate useful information about:

* Total Copies
* Books Borrowed
* Available Copies
* Borrowing Rate
* Department-wise borrowing
* Most borrowed book
* Least borrowed book
* Books with no available copies
* Relationship between total copies and books borrowed
* Data visualization

---

# 2. Complete Project Workflow

The project follows these **8 stages**:

```text
Data Collection
       ↓
Data Entry
       ↓
Data Validation
       ↓
Data Storage
       ↓
Data Processing
       ↓
Data Analysis
       ↓
Reporting
       ↓
Visualization
```

---

# 3. Raw Data

The following raw data was collected from different sources:

```text
B001, Introduction to Algorithms, CSE, Thomas Cormen, 2019, 5, 12
B002, Database System Concepts, CSE, Abraham Silberschatz, 2018, 4, 10
B003, Computer Networks, CSE, Andrew Tanenbaum, 2020, 3, 8
B004, Operating System Concepts, CSE, Abraham Silberschatz, 2018, 6, 15
B005, Engineering Mathematics, EEE, Erwin Kreyszig, 2017, 5, 7
B006, Digital Logic Design, EEE, Morris Mano, 2019, 4, 9
B007, Principles of Marketing, BBA, Philip Kotler, 2021, 3, 11
B008, Financial Accounting, BBA, Jerry Weygandt, 2020, 4, 6
B009, Software Engineering, CSE, Ian Sommerville, 2019, 5, 13
B010, Microprocessor Architecture, EEE, Ramesh Gaonkar, 2018, 2, 5
B011, Data Structures, CSE, Seymour Lipschutz, 2020, 6, 14
B012, Business Communication, BBA, Mary Guffey, 2021, 3, 8
B013, Artificial Intelligence, CSE, Stuart Russell, 2020, 4, 12
B014, Electronic Devices, EEE, Thomas Floyd, 2019, 5, 10
B015, Human Resource Management, BBA, Gary Dessler, 2021, 2, 4
```

---

# 4. Data Fields

The raw data contains the following fields:

| Column | Field            |
| ------ | ---------------- |
| A      | Book ID          |
| B      | Book Title       |
| C      | Department       |
| D      | Author           |
| E      | Publication Year |
| F      | Total Copies     |
| G      | Books Borrowed   |

Two new fields will be added later:

| Column | New Field        |
| ------ | ---------------- |
| H      | Available Copies |
| I      | Borrowing Rate   |

---

# 5. Stage 1 — Data Collection

## Task

Identify at least **3 possible sources** from which library book data could be collected.

## Possible Sources

### 1. Library Management System

The library management system can provide:

* Book ID
* Book Title
* Author
* Department
* Publication Year
* Total Copies

### 2. Book Issue / Borrowing Records

The borrowing system can provide:

* Books Borrowed
* Borrowing history
* Issue records

### 3. University Department Records

Academic departments can provide:

* Department-wise book information
* Recommended textbooks
* Course-related books
* Book categories

### 4. Library Inventory Records

Inventory records can provide:

* Total copies
* Available copies
* Damaged or missing books
* New book additions

---

## Source Mapping

| Data             | Possible Source              |
| ---------------- | ---------------------------- |
| Book ID          | Library Management System    |
| Book Title       | Library Management System    |
| Department       | Department / Library Records |
| Author           | Library Catalog              |
| Publication Year | Library Catalog              |
| Total Copies     | Library Inventory            |
| Books Borrowed   | Borrowing / Issue Records    |

---

# 6. Stage 2 — Data Entry

Open Microsoft Excel.

---

## Step 1 — Create a New Workbook

Open:

```text
Microsoft Excel
→ Blank Workbook
```

---

## Step 2 — Save the File

Go to:

```text
File
→ Save As
```

Use a suitable filename:

```text
University_Library_Book_Data_Processing.xlsx
```

---

## Step 3 — Enter Headers

Enter the following headers in Row 1:

| Cell | Header           |
| ---- | ---------------- |
| A1   | Book ID          |
| B1   | Book Title       |
| C1   | Department       |
| D1   | Author           |
| E1   | Publication Year |
| F1   | Total Copies     |
| G1   | Books Borrowed   |

---

## Step 4 — Enter the Raw Data

Enter the 15 records from:

```text
Row 2 → Row 16
```

Therefore:

```text
A1:G16
```

will contain the complete raw dataset.

---

## Final Initial Structure

```text
A = Book ID
B = Book Title
C = Department
D = Author
E = Publication Year
F = Total Copies
G = Books Borrowed
```

---


# 7. Stage 3 — Data Validation

The purpose of data validation is to identify incorrect, missing, duplicated, or inconsistent data.

We will check:

```text
1. Missing Data
2. Duplicate Book IDs
3. Publication Year
4. Total Copies
5. Books Borrowed
6. Borrowed ≤ Total Copies
7. Department
```

---

# 7.1 Missing Data Check

Select:

```text
A2:G16
```

Then:

```text
Home
→ Find & Select
→ Go To Special
→ Blanks
→ OK
```

### Expected Result

If no blank cells are selected:

```text
No missing data found.
```

For this dataset:

**No missing values were found.**

---

# 7.2 Duplicate Book ID Check

Select:

```text
A2:A16
```

Then:

```text
Home
→ Conditional Formatting
→ Highlight Cells Rules
→ Duplicate Values
```

Click:

```text
OK
```

### Expected Result

No Book IDs should be highlighted.

```text
No duplicate Book IDs found.
```

The dataset contains unique Book IDs:

```text
B001 → B015
```

---

# 7.3 Publication Year Validation

Publication Year is stored in:

```text
E2:E16
```

The given dataset contains publication years:

```text
2017
2018
2019
2020
2021
```

These are valid years for the dataset.

---

## Excel Check

Select:

```text
E2:E16
```

Then:

```text
Home
→ Conditional Formatting
→ Highlight Cells Rules
→ Less Than
```

Enter:

```text
1900
```

Then also check values above the current year if required.

For this assignment, the given years are valid.

### Result

```text
All publication years are valid.
```

---

# 7.4 Total Copies Validation

Total Copies are stored in:

```text
F2:F16
```

Total copies should be positive numbers.

Select:

```text
F2:F16
```

Then:

```text
Home
→ Conditional Formatting
→ Highlight Cells Rules
→ Less Than
```

Enter:

```text
1
```

If no values are highlighted:

```text
All Total Copies values are valid.
```

---

# 7.5 Books Borrowed Validation

Books Borrowed are stored in:

```text
G2:G16
```

Books Borrowed should be numeric and should not be negative.

Select:

```text
G2:G16
```

Then:

```text
Home
→ Conditional Formatting
→ Highlight Cells Rules
→ Less Than
```

Enter:

```text
0
```

No negative values are present in the dataset.

---

# 7.6 Check Whether Borrowed Books Exceed Total Copies

This is the most important validation in this dataset.

The requirement says:

```text
Books Borrowed ≤ Total Copies
```

We can check this using an additional temporary column.

For example, enter in:

```text
J1
```

```text
Validation Status
```

Then in:

```text
J2
```

enter:

```excel
=IF(G2>F2,"Invalid","Valid")
```

Press:

```text
Enter
```

Then double-click the Fill Handle.

The formula will fill:

```text
J2:J16
```

---

## Formula Explanation

```excel
=IF(G2>F2,"Invalid","Valid")
```

means:

```text
If Books Borrowed > Total Copies
        ↓
      Invalid

Otherwise
        ↓
       Valid
```

---

# 7.7 Validation Result

The provided dataset contains a major inconsistency.

For example:

```text
B001
Total Copies = 5
Books Borrowed = 12
```

But:

```text
12 > 5
```

Therefore:

```text
B001 = Invalid
```

The same issue occurs for the other records as well.

### Examples

| Book ID | Total Copies | Books Borrowed | Status  |
| ------- | -----------: | -------------: | ------- |
| B001    |            5 |             12 | Invalid |
| B002    |            4 |             10 | Invalid |
| B003    |            3 |              8 | Invalid |
| B004    |            6 |             15 | Invalid |
| B005    |            5 |              7 | Invalid |
| B006    |            4 |              9 | Invalid |
| B007    |            3 |             11 | Invalid |
| B008    |            4 |              6 | Invalid |
| B009    |            5 |             13 | Invalid |
| B010    |            2 |              5 | Invalid |
| B011    |            6 |             14 | Invalid |
| B012    |            3 |              8 | Invalid |
| B013    |            4 |             12 | Invalid |
| B014    |            5 |             10 | Invalid |
| B015    |            2 |              4 | Invalid |

Therefore:

```text
15 out of 15 records violate the condition
Books Borrowed ≤ Total Copies.
```

---

# 7.8 Department Validation

The departments in this dataset are:

```text
CSE
EEE
BBA
```

Select:

```text
C2:C16
```

You can use:

```text
Data
→ Data Validation
```

and create a list:

```text
CSE,EEE,BBA
```

This prevents invalid department names from being entered.

For the current dataset:

```text
All department values are valid.
```

---