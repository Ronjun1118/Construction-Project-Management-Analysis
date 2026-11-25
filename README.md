# Construction Project Management Analysis

## Project Overview
This project focuses on analyzing construction project management data using Microsoft Excel to identify operational trends, monitor project performance, and support data-driven decision-making.

## Project Objectives
- Clean and validate construction project data.
- Identify duplicate records, missing values, and data inconsistencies.
- Analyze open, closed, and overdue forms.
- Track outstanding actions across construction projects.
- Identify monthly reporting trends.
- Develop an interactive dashboard to visualize key performance indicators (KPIs).

## Tools and Skills
- Microsoft Excel
- Data Cleaning and Validation
- Excel Formulas and Functions
- Conditional Formatting
- Pivot Tables
- Charts and Data Visualization
- Dashboard Development
- KPI Reporting

## Data Analysis Process
1. Ask - Define business questions.
2. Prepare - Collect and understand the dataset.
3. Process - Clean and validate data.
4. Analyze - Identify patterns and trends.
5. Share - Present findings through dashboards.
6. Act - Provide data-driven recommendations.

## Data Cleaning Progress
- [ ] Check duplicate records
- [ ] Identify missing values
- [ ] Remove unnecessary spaces
- [ ] Standardize data formats
- [ ] Validate date formats
- [ ] Check data consistency

## Project Status
## Phase 1: Ask – Business Understanding

**Status:** Completed

### Business Problem

Construction project management involves tracking multiple projects, inspection forms, outstanding actions, and project statuses. Without effective data analysis, identifying overdue activities, unresolved issues, and project performance trends can be challenging.

This project aims to analyze construction project management data to improve visibility into project performance, identify operational inefficiencies, and support data-driven decision-making.

### Business Objectives

- Monitor overall construction project performance.
- Identify open, closed, and overdue forms.
- Evaluate outstanding actions across construction projects.
- Identify projects with the highest number of unresolved actions.
- Analyze construction reporting trends over time.
- Develop an interactive Excel dashboard to communicate key performance indicators.
- Provide actionable recommendations to improve project monitoring and operational efficiency.

### Key Business Questions

1. How many construction projects and report forms are included in the dataset?
2. What percentage of report forms are open and closed?
3. How many outstanding actions require attention?
4. Which projects have the highest number of open actions?
5. Which projects have the highest number of overdue forms?
6. Which report form types have the most outstanding actions?
7. How does construction reporting activity change over time?
8. What opportunities exist to improve project performance and operational efficiency?

### Target Stakeholders

- Project Managers
- Project Coordinators
- Construction Managers
- Operations Managers
- Quality Assurance Teams
- Business Analysts

### Expected Deliverables

- Cleaned and validated Excel dataset
- Data cleaning documentation
- Key Performance Indicators (KPIs)
- Pivot tables and charts
- Interactive Excel dashboard
- Key findings and actionable business recommendations

- ## Phase 2: Prepare – Data Collection and Understanding

**Status:** Completed

### Dataset Overview

The dataset contains construction project management records, including report forms, project names, status information, action counts, overdue indicators, and reporting dates.

| Dataset Information | Description |
|---|---|
| Dataset Name | Construction Project Management |
| Data Source | Add original dataset source and link |
| Original File Format | CSV |
| Analysis Tool | Microsoft Excel |
| Original Record Count | 10,254 |
| Original Column Count | 17 |
| Data Type | Structured, tabular data |
| Industry | Construction and Project Management |

### Data Dictionary

| Column | Field Name | Description |
|---|---|---|
| A | Ref | Report form reference identifier |
| B | Status | Detailed report status |
| C | Location | Location associated with the record |
| D | Name | Name or title associated with the report |
| E | Created | Date the report was created |
| F | Type | Report or form type |
| G | Status Changed | Date the report status was changed |
| H | Open Actions | Number of outstanding actions |
| I | Total Actions | Total number of associated actions |
| J | Association | Association information related to the record |
| K | OverDue | Overdue status indicator |
| L | Images | Image-related information |
| M | Comments | Comment-related information |
| N | Documents | Document-related information |
| O | Project | Construction project name |
| P | Report Forms Status | Overall report form status |
| Q | Report Forms Group | Report form classification or group |

### Data Preparation Activities

- Imported the original CSV dataset into Microsoft Excel.
- Reviewed the dataset structure, column names, and record count.
- Examined numerical, categorical, and date fields.
- Identified relevant fields required to answer business questions.
- Preserved the original dataset to maintain data integrity.
- Created a separate worksheet for data cleaning and transformation.
- Created a data cleaning log to document validation results and corrective actions.

### Excel Workbook Structure

| Worksheet | Purpose |
|---|---|
| Raw_Data | Preserve the original dataset |
| Cleaned_Data | Store cleaned and validated records |
| Data_Cleaning_Log | Document data quality checks and corrective actions |
| Analysis | Perform exploratory data analysis and develop KPIs |
| Dashboard | Present key metrics and interactive visualizations |

### Data Quality Considerations

The following potential data quality issues were considered before analysis:

- Duplicate records
- Missing values
- Unnecessary spaces
- Inconsistent date formats
- Invalid dates
- Inconsistent action counts
- Unavailable status values
- Missing project names
- Record count discrepancies

### Data Preparation Outcome

The dataset was imported, organized, and reviewed to establish a structured foundation for data cleaning, exploratory data analysis, and dashboard development.

## Phase 3: Process – Data Cleaning and Validation

**Status:** Completed

Data cleaning was performed using Microsoft Excel to ensure data accuracy, consistency, completeness, and reliability before conducting exploratory data analysis.

The original dataset was preserved in the `Raw_Data` worksheet, while all data cleaning activities were performed in the `Cleaned_Data` worksheet. A separate `Data_Cleaning_Log` worksheet was maintained to document cleaning activities, results, and corrective actions.

### 1. Duplicate Records

- Reviewed the dataset for duplicate records.
- Used Excel's Remove Duplicates feature to identify and remove identical records.
- Reviewed repeated reference identifiers to avoid removing legitimate records.

**Excel Formula:**

`=IF(COUNTIF($A$2:$A$10255,A2)>1,"Duplicate","Unique")`

### 2. Missing Values

- Identified missing values across relevant data fields.
- Reviewed missing information before determining appropriate corrective actions.
- Preserved records containing useful information to prevent unnecessary data loss.

**Excel Formula:**

`=COUNTBLANK(Cleaned_Data!A2:A10255)`

### 3. Removal of Unnecessary Spaces

- Reviewed text fields for leading, trailing, and unnecessary spaces.
- Used the TRIM function to standardize text values and improve data consistency.

**Excel Formulas:**

`=TRIM(B2)`

`=SUMPRODUCT(--(Cleaned_Data!B2:B10255<>TRIM(Cleaned_Data!B2:B10255)))`

### 4. Date Format Standardization

- Converted text-formatted dates into valid Excel dates using the Text to Columns feature.
- Standardized the Created and Status Changed date fields.
- Used the ISNUMBER function to verify date conversion.

**Excel Formulas:**

`=ISNUMBER(E2)`

`=ISNUMBER(G2)`

### 5. Date Validation

- Validated Created and Status Changed dates.
- Checked for missing dates, invalid formats, and chronological inconsistencies.
- Confirmed that all records passed the applied date validation checks.

**Excel Formula:**

`=IF(OR(E2="",G2=""),"Missing Date",IF(OR(NOT(ISNUMBER(E2)),NOT(ISNUMBER(G2))),"Invalid Format",IF(G2<E2,"Invalid Date","Valid")))`

### 6. Data Consistency Validation

#### Action Count Validation

- Validated Open Actions and Total Actions.
- Checked for missing, nonnumeric, negative, and inconsistent action counts.
- Verified whether Open Actions exceeded Total Actions.

**Excel Formula:**

`=IF(OR(H2="",I2=""),"Missing Value",IF(OR(NOT(ISNUMBER(H2)),NOT(ISNUMBER(I2))),"Invalid Format",IF(OR(H2<0,I2<0,H2>I2),"Inconsistent","Valid")))`

#### Status Consistency

- Reviewed unique status categories.
- Identified Open, Closed, and unavailable status values.
- Flagged unavailable values to prevent incorrect status classification.

**Excel Formulas:**

`=SORT(UNIQUE(Cleaned_Data!B2:B10255))`

`=SORT(UNIQUE(Cleaned_Data!P2:P10255))`

`=SUMPRODUCT(--ISNA(Cleaned_Data!P2:P10255))`

#### Overdue Indicator Validation

- Reviewed unique overdue indicator values to identify inconsistent categories.

**Excel Formula:**

`=UNIQUE(Cleaned_Data!K2:K10255)`

#### Project Name Consistency

- Reviewed unique project names.
- Identified missing project information.
- Reviewed project naming consistency to support accurate project-level reporting.

**Excel Formulas:**

`=SORT(UNIQUE(Cleaned_Data!O2:O10255))`

`=COUNTBLANK(Cleaned_Data!O2:O10255)`

### 7. Final Record Count Verification

- Compared original and cleaned dataset record counts.
- Verified the number of records removed during data cleaning.
- Reviewed record counts to prevent unintended data loss.

**Original Record Count:**

`=COUNTA(Raw_Data!A:A)-1`

**Final Record Count:**

`=COUNTA(Cleaned_Data!A:A)-1`

**Records Removed:**

`=COUNTA(Raw_Data!A:A)-COUNTA(Cleaned_Data!A:A)`

### Data Cleaning Summary

| Metric | Result |
|---|---|
| Original Records | 10,254 |
| Original Columns | 17 |
| Duplicate Records Removed | Enter actual count |
| Final Record Count | Enter actual count |
| Date Validation Errors | 0 |
| Data Cleaning Status | Completed |

### Excel Functions and Tools Used

- IF
- OR
- NOT
- COUNTIF
- COUNTA
- COUNTBLANK
- SUMPRODUCT
- TRIM
- ISNUMBER
- ISNA
- SORT
- UNIQUE
- Remove Duplicates
- Text to Columns
- Data Filtering
- Data Validation Checks

### Data Cleaning Outcome

The dataset was reviewed and processed using Excel data cleaning and validation techniques to prepare it for exploratory data analysis, KPI development, pivot tables, and dashboard creation.


## Dashboard
Coming soon.

## Key Findings
To be updated after completing the analysis.

## Recommendations
To be updated based on analytical findings.
