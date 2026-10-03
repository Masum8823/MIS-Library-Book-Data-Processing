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

# 8. Stage 4 — Data Storage

After data validation and cleaning, the dataset should be stored in a structured digital format.

## Recommended Storage Method

Use an **Excel `.xlsx` file**.

Example:

```text
University_Library_Book_Data_Processing.xlsx
```

---

## Why Excel?

Excel is suitable because it provides:

* Structured rows and columns
* Easy data entry
* Formulas
* Sorting and filtering
* Conditional Formatting
* Data Validation
* Charts
* Easy sharing
* Easy reporting

---

## Optional CSV Storage

The cleaned dataset can also be exported as:

```text
University_Library_Book_Data.csv
```

CSV is useful when the data needs to be imported into:

* Database systems
* Python
* R
* Other data analysis software

---

# 9. Stage 5 — Data Processing

Now we calculate two new fields:

```text
Available Copies
Borrowing Rate
```

---

# 9.1 Add Available Copies Column

Enter:

```text
H1
```

```text
Available Copies
```

Formula:

```text
Available Copies
=
Total Copies − Books Borrowed
```

---

## Excel Formula

Click:

```text
H2
```

Enter:

```excel
=F2-G2
```

Press:

```text
Enter
```

Then double-click the Fill Handle.

The formula will automatically fill:

```text
H2:H16
```

---

## Formula Explanation

```text
F2 = Total Copies
G2 = Books Borrowed
```

Therefore:

```excel
=F2-G2
```

means:

```text
Total Copies − Books Borrowed
```

---

# 9.2 Add Borrowing Rate Column

Enter:

```text
I1
```

```text
Borrowing Rate
```

The formula is:

```text
Borrowing Rate
=
(Books Borrowed ÷ Total Copies) × 100
```

---

## Excel Formula

Click:

```text
I2
```

Enter:

```excel
=G2/F2*100
```

Press:

```text
Enter
```

Then double-click the Fill Handle.

The formula will fill:

```text
I2:I16
```

---

# 9.3 Example Calculation

For B001:

```text
Total Copies = 5
Books Borrowed = 12
```

Available Copies:

```text
5 − 12 = -7
```

Borrowing Rate:

```text
(12 ÷ 5) × 100
= 240%
```

So:

```text
B001 Available Copies = -7
B001 Borrowing Rate = 240%
```

---

# 9.4 Important Note About the Dataset

Normally, available copies should not be negative.

Also, if Books Borrowed represents the current number of borrowed copies, borrowing rate should normally not exceed 100%.

However, the supplied raw data contains values where:

```text
Books Borrowed > Total Copies
```

Therefore, the formulas produce negative Available Copies and Borrowing Rates above 100%.

This is useful for demonstrating the **Data Validation** stage.

---

# 10. Stage 6 — Data Analysis

Now we analyze the processed data.

---

# 10.1 Total Number of Copies

Formula:

```excel
=SUM(F2:F16)
```

Result:

```text
61
```

Therefore:

```text
Total Copies = 61
```

---

# 10.2 Total Books Borrowed

Formula:

```excel
=SUM(G2:G16)
```

Result:

```text
144
```

Therefore:

```text
Total Books Borrowed = 144
```

---

# 10.3 Total Available Copies

Formula:

```excel
=SUM(H2:H16)
```

Result:

```text
-83
```

This negative result occurs because the raw data contains invalid records where borrowed books exceed total copies.

Therefore:

```text
Calculated Available Copies = -83
```

### Important

This should **not** be interpreted as 83 physically negative books. It is an indicator that the supplied dataset is inconsistent.

---

# 10.4 Average Borrowing Rate

Formula:

```excel
=AVERAGE(I2:I16)
```

Result:

```text
239.89%
```

Therefore:

```text
Average Borrowing Rate ≈ 239.89%
```

Again, this unusually high value is caused by the invalid raw data.

---

# 10.5 Highest Borrowing Rate

Formula:

```excel
=MAX(I2:I16)
```

Result:

```text
366.67%
```

This belongs to:

```text
B007 — Principles of Marketing
```

Calculation:

```text
11 ÷ 3 × 100
= 366.67%
```

---

# 10.6 Lowest Borrowing Rate

Formula:

```excel
=MIN(I2:I16)
```

Result:

```text
140%
```

This belongs to:

```text
B005 — Engineering Mathematics
```

Calculation:

```text
7 ÷ 5 × 100
= 140%
```

---

# 10.7 Department-wise Average Borrowing Rate

Create a separate table.

For example:

| K          | L                      |
| ---------- | ---------------------- |
| Department | Average Borrowing Rate |
| CSE        |                        |
| EEE        |                        |
| BBA        |                        |

---

## CSE

In L2:

```excel
=AVERAGEIF(C2:C16,"CSE",I2:I16)
```

Result:

```text
257.14%
```

---

## EEE

In L3:

```excel
=AVERAGEIF(C2:C16,"EEE",I2:I16)
```

Result:

```text
203.75%
```

---

## BBA

In L4:

```excel
=AVERAGEIF(C2:C16,"BBA",I2:I16)
```

Result:

```text
245.83%
```

---

## Final Table

| Department | Average Borrowing Rate |
| ---------- | ---------------------: |
| CSE        |                257.14% |
| EEE        |                203.75% |
| BBA        |                245.83% |

---

# 10.8 Department-wise Total Books Borrowed

Create another table:

| K          | L                    |
| ---------- | -------------------- |
| Department | Total Books Borrowed |
| CSE        |                      |
| EEE        |                      |
| BBA        |                      |

---

## CSE

```excel
=SUMIF(C2:C16,"CSE",G2:G16)
```

Result:

```text
84
```

---

## EEE

```excel
=SUMIF(C2:C16,"EEE",G2:G16)
```

Result:

```text
31
```

---

## BBA

```excel
=SUMIF(C2:C16,"BBA",G2:G16)
```

Result:

```text
29
```

---

## Final Table

| Department | Total Books Borrowed |
| ---------- | -------------------: |
| CSE        |                   84 |
| EEE        |                   31 |
| BBA        |                   29 |

---

# 10.9 Most Borrowed Book

We need to find the highest value in:

```text
G2:G16
```

Use:

```excel
=MAX(G2:G16)
```

Result:

```text
15
```

Now find which book has 15 borrowed books.

From the dataset:

```text
B004 — Operating System Concepts
```

Therefore:

```text
Most Borrowed Book = Operating System Concepts
Books Borrowed = 15
```

---

## Optional Formula to Find Book Title

If your Excel supports `XLOOKUP`:

```excel
=XLOOKUP(MAX(G2:G16),G2:G16,B2:B16)
```

Result:

```text
Operating System Concepts
```

---

## Alternative Formula

If `XLOOKUP` is unavailable:

```excel
=INDEX(B2:B16,MATCH(MAX(G2:G16),G2:G16,0))
```

---

# 10.10 Least Borrowed Book

Use:

```excel
=MIN(G2:G16)
```

Result:

```text
4
```

The book with 4 borrowed copies is:

```text
B015 — Human Resource Management
```

Therefore:

```text
Least Borrowed Book = Human Resource Management
Books Borrowed = 4
```

---

## Optional Formula

```excel
=XLOOKUP(MIN(G2:G16),G2:G16,B2:B16)
```

Result:

```text
Human Resource Management
```

---

# 10.11 Number of Books with No Available Copies

The task asks for:

```text
Number of books with no available copies
```

Normally we would use:

```excel
=COUNTIF(H2:H16,0)
```

However, because this dataset produces **negative Available Copies**, there are no records exactly equal to zero.

Result:

```text
0
```

But this is another indication of the data inconsistency.

---

## Better Validation Check

To identify records where no copies are available **or the calculation becomes invalid**, use:

```excel
=COUNTIF(H2:H16,"<=0")
```

Result:

```text
15
```

Therefore:

```text
15 records have zero or negative calculated availability.
```

This is consistent with the validation result that every record has:

```text
Books Borrowed > Total Copies
```

---

# 10.12 Relationship Between Total Copies and Books Borrowed

Use the `CORREL()` function.

Total Copies:

```text
F2:F16
```

Books Borrowed:

```text
G2:G16
```

Formula:

```excel
=CORREL(F2:F16,G2:G16)
```

Result:

```text
Approximately 0.770
```

---

## Interpretation

The correlation coefficient is approximately:

```text
r = 0.770
```

This indicates a **positive relationship** between Total Copies and Books Borrowed in this dataset.

In simple terms:

> Books with more total copies generally tend to have more borrowing activity in the given dataset.

---

# 11. Stage 7 — Reporting

The major findings can be summarized in a report.

## Library Data Processing Summary

The dataset contains:

```text
15 books
61 total copies
144 recorded borrowed books
```

The calculated average borrowing rate is:

```text
239.89%
```

The highest borrowing rate is:

```text
366.67%
```

for:

```text
Principles of Marketing
```

The lowest borrowing rate is:

```text
140%
```

for:

```text
Engineering Mathematics
```

Department-wise total borrowing:

```text
CSE = 84
EEE = 31
BBA = 29
```

The most borrowed book is:

```text
Operating System Concepts
15 borrowed
```

The least borrowed book is:

```text
Human Resource Management
4 borrowed
```

The correlation between total copies and books borrowed is approximately:

```text
0.770
```

indicating a positive relationship.

---

# 12. Stage 8 — Data Visualization

The project requires at least 2 charts.

We will create **5 useful charts**:

```text
1. Department-wise Books Borrowed
2. Book-wise Borrowing Rate
3. Available vs Borrowed Copies
4. Department-wise Total Copies
5. Total Copies vs Books Borrowed
```

---

# 12.1 Chart 1 — Department-wise Books Borrowed

First create this table:

| Department | Total Books Borrowed |
| ---------- | -------------------: |
| CSE        |                   84 |
| EEE        |                   31 |
| BBA        |                   29 |

---

## Excel Steps

Select:

```text
K1:L4
```

Then:

```text
Insert
→ Column or Bar Chart
→ 2-D Column
→ Clustered Column
```

Change the chart title to:

```text
Department-wise Books Borrowed
```

### Axis

```text
X-axis = Department
Y-axis = Total Books Borrowed
```

---

# 12.2 Chart 2 — Book-wise Borrowing Rate

We need:

```text
Book Title
Borrowing Rate
```

These are:

```text
B2:B16
I2:I16
```

Because the two columns are separated, you can create a small helper table.

For example:

| K                          | L              |
| -------------------------- | -------------- |
| Book Title                 | Borrowing Rate |
| Introduction to Algorithms | 240%           |
| Database System Concepts   | 250%           |
| Computer Networks          | 266.67%        |
| ...                        | ...            |

Then select the table.

Go to:

```text
Insert
→ Column Chart
→ 2-D Column
→ Clustered Column
```

Chart title:

```text
Book-wise Borrowing Rate
```

---

# 12.3 Chart 3 — Available vs Borrowed Copies

Create a helper table:

| Book Title                 | Available Copies | Books Borrowed |
| -------------------------- | ---------------: | -------------: |
| Introduction to Algorithms |               -7 |             12 |
| Database System Concepts   |               -6 |             10 |
| Computer Networks          |               -5 |              8 |
| ...                        |              ... |            ... |

Select the table.

Then:

```text
Insert
→ Column Chart
→ Clustered Column
```

Chart title:

```text
Available vs. Borrowed Copies
```

### Axis

```text
X-axis = Book Title
Y-axis = Number of Copies
```

### Important

Because the raw dataset is inconsistent, Available Copies will contain negative values. The chart will visually show this problem.

---

# 12.4 Chart 4 — Department-wise Total Copies

Create:

| Department | Total Copies |
| ---------- | -----------: |
| CSE        |           33 |
| EEE        |           16 |
| BBA        |           12 |

---

## Formulas

### CSE

```excel
=SUMIF(C2:C16,"CSE",F2:F16)
```

Result:

```text
33
```

### EEE

```excel
=SUMIF(C2:C16,"EEE",F2:F16)
```

Result:

```text
16
```

### BBA

```excel
=SUMIF(C2:C16,"BBA",F2:F16)
```

Result:

```text
12
```

---

## Create Chart

Select the table:

```text
K8:L11
```

Then:

```text
Insert
→ Column Chart
→ Clustered Column
```

Title:

```text
Department-wise Total Copies
```

---

# 12.5 Chart 5 — Total Copies vs Books Borrowed

This chart uses a **Scatter Chart**.

Required data:

```text
X-axis = Total Copies
Y-axis = Books Borrowed
```

Select:

```text
F1:G16
```

Then:

```text
Insert
→ Scatter (X, Y)
→ Scatter with only Markers
```

Change the title to:

```text
Total Copies vs. Books Borrowed
```

---

## Add Trendline

Click the chart.

Then:

```text
Chart
→ +
→ Trendline
→ Linear
```

The trendline will show the general relationship between:

```text
Total Copies
```

and:

```text
Books Borrowed
```

The correlation value:

```text
r ≈ 0.770
```

shows a positive relationship in the given data.

---

# 13. Important Data Validation Finding

This is an important part of the project.

The raw dataset contains a logical inconsistency.

The requirement says:

```text
Books Borrowed ≤ Total Copies
```

But the provided data contains:

```text
Books Borrowed > Total Copies
```

for every record.

For example:

```text
B001

Total Copies = 5
Books Borrowed = 12
```

This gives:

```text
12 > 5
```

which is invalid if `Books Borrowed` means the **current number of copies physically borrowed at one time**.

---

## Effect on Calculations

Because of this issue:

```text
Available Copies = Total Copies − Books Borrowed
```

produces negative values.

Example:

```text
5 − 12 = -7
```

Similarly:

```text
Borrowing Rate
= 12 / 5 × 100
= 240%
```

Therefore, the calculated:

```text
Available Copies = -83
Average Borrowing Rate ≈ 239.89%
```

should not be interpreted as real-world library inventory values.

---

## Possible Explanation

One possible interpretation is that **Books Borrowed** represents the total number of borrowing transactions over a period rather than the number of copies currently checked out.

Under that interpretation, a book can be borrowed multiple times, so the borrowing count can exceed the number of physical copies.

However, this interpretation conflicts with the specific validation requirement:

```text
Books Borrowed do not exceed Total Copies
```

Therefore, for this assignment, the correct approach is to **identify and report the inconsistency during Data Validation**.

---

# 14. Final Analysis Summary

| Analysis                                 |                         Result |
| ---------------------------------------- | -----------------------------: |
| Number of Book Records                   |                             15 |
| Total Copies                             |                             61 |
| Total Books Borrowed                     |                            144 |
| Calculated Available Copies              |                            -83 |
| Average Borrowing Rate                   |                        239.89% |
| Highest Borrowing Rate                   |                        366.67% |
| Lowest Borrowing Rate                    |                           140% |
| CSE Average Borrowing Rate               |                        257.14% |
| EEE Average Borrowing Rate               |                        203.75% |
| BBA Average Borrowing Rate               |                        245.83% |
| CSE Total Borrowed                       |                             84 |
| EEE Total Borrowed                       |                             31 |
| BBA Total Borrowed                       |                             29 |
| Most Borrowed Book                       | Operating System Concepts — 15 |
| Least Borrowed Book                      |  Human Resource Management — 4 |
| Books with Exactly 0 Available Copies    |                              0 |
| Records with ≤ 0 Calculated Availability |                             15 |
| Total Copies vs Borrowed Correlation     |                        ≈ 0.770 |
| Borrowed > Total Copies                  |                     15 records |

---

# 15. Final Excel Structure

After processing, the main dataset should look like this:

| Column | Header           |
| ------ | ---------------- |
| A      | Book ID          |
| B      | Book Title       |
| C      | Department       |
| D      | Author           |
| E      | Publication Year |
| F      | Total Copies     |
| G      | Books Borrowed   |
| H      | Available Copies |
| I      | Borrowing Rate   |

---

## Example

| Book ID | Book Title                 | Dept. | Year | Total | Borrowed | Available |    Rate |
| ------- | -------------------------- | ----- | ---: | ----: | -------: | --------: | ------: |
| B001    | Introduction to Algorithms | CSE   | 2019 |     5 |       12 |        -7 |    240% |
| B002    | Database System Concepts   | CSE   | 2018 |     4 |       10 |        -6 |    250% |
| B003    | Computer Networks          | CSE   | 2020 |     3 |        8 |        -5 | 266.67% |
| B004    | Operating System Concepts  | CSE   | 2018 |     6 |       15 |        -9 |    250% |
| ...     | ...                        | ...   |  ... |   ... |      ... |       ... |     ... |
| B015    | Human Resource Management  | BBA   | 2021 |     2 |        4 |        -2 |    200% |

---
