\# Excel Automation System using Power Query



\## Overview



This project demonstrates a fully automated Excel reporting system built entirely using Power Query.



The system consolidates multiple monthly Excel files, applies business logic, performs automated data mapping, detects unmapped records, generates management-ready reports, and updates everything with a single refresh.



It is designed for real-world business environments where data is updated regularly and manual processing is time-consuming.



---



\## Problem Statement



Many companies manage monthly operational data using separate Excel files.



Common challenges include:

\- Manual file consolidation

\- Repetitive VLOOKUP / XLOOKUP tasks

\- Inconsistent column formats

\- Missing mapping records

\- Manual summary report creation

\- Risk of formula breakage



This system eliminates all of the above using Power Query.



---



\## Key Features



\### 1. Folder-Based Automatic Consolidation

\- Automatically loads multiple Excel files from a folder

\- Handles inconsistent column names

\- Supports incremental monthly file additions

\- Updates with a single refresh



\### 2. Automated Mapping System (Merge Logic)

\- Branch + Item → SalesManager

\- Item → Category \& Account

\- Branch → Region \& CostCenter

\- Fully dynamic and refresh-safe



\### 3. UNMAPPED Detection (Data Governance Layer)

\- Automatically detects new Branch/Item combinations

\- Generates a clean list for mapping table updates

\- Prevents silent reporting errors



\### 4. Reporting Layer

\- Monthly summary

\- Branch by month performance

\- Category summary

\- Top N products

\- All reports update automatically



\### 5. Error Monitoring

\- Detects corrupted or invalid Excel files

\- Prevents refresh failure

\- Separates problematic files from clean data



---



\## System Architecture



Raw Excel Files (Monthly)

&nbsp;       ↓

Power Query (ETL Engine)

&nbsp;       ↓

Mapping Layer (Merge)

&nbsp;       ↓

UNMAPPED Detection

&nbsp;       ↓

Reporting Layer

&nbsp;       ↓

Final Report



---



\## Workflow



1\. Add a new monthly file to the folder

2\. Click Refresh

3\. Mapping logic is applied automatically

4\. Unmapped records are identified

5\. Reports are regenerated



No manual data processing required.



---



\## Technical Stack



\- Microsoft Excel

\- Power Query (M language)

\- Folder-based ETL logic

\- Left Anti Join for unmapped detection

\- Group By for reporting aggregation



---



\## Business Value



\- Eliminates repetitive manual work

\- Reduces human error

\- Improves reporting consistency

\- Enables scalable monthly operations

\- Maintains data governance visibility

\- Non-technical user friendly



---



\## Demo Video



1.https://youtu.be/RrSQbbsh-Ys



2.https://youtu.be/TDYwmUnXW\_M



3.https://youtu.be/2vgHFhHgGe4



4.https://youtu.be/7\_V0wrHN49Q



5.https://youtu.be/8wNIxj2iOZg



6.https://youtu.be/cLSN4DWRDIQ



7.https://youtu.be/8fiUnF1kvlI



