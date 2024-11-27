# Talent-Port Data Analytics
## Excel 
---
### Interface and Navigation

Tabs/Ribbons: Excel's interface is organized into various tabs and ribbons, providing access to different functionalities and commands.
Cell Navigation: Use the arrow keys to move between cells (up, down, left, right). Press Enter to move down to the next row, and Shift + Enter to move up to the previous row.

### Typing and Editing in Excel

Wrapping Text: Use Alt + Enter to wrap text within a cell, allowing you to display multiple lines of text.
In-Cell Editing: Press F2 to enter in-cell editing mode, which allows you to directly edit the contents of a cell without overwriting the entire cell.

### Data Manipulation

Text to Columns: The "Text to Columns" feature, found in the Data ribbon, allows you to split text data into multiple columns based on a specified delimiter (e.g., comma, space, tab).
Delimiters: Delimiters are the characters used to separate data, such as commas, tabs, or spaces. Correctly identifying the delimiter is important when using the "Text to Columns" feature.

class Activity: Typed into cells, Converted text to clomuns, Changed text and character colors (flash fill)

# Week 2 Excel

## Table of Contents
1. [Cell Formatting and Manipulation](#cell-formatting-and-manipulation)
2. [Copy and Paste Operations](#copy-and-paste-operations)
3. [Formula Fundamentals](#formula-fundamentals)
4. [Pivot Tables](#pivot-tables)

## Cell Formatting and Manipulation

### Cell Selection and Formatting
1. **Basic Selection**
   - Click and drag to select multiple cells
   - Use Shift + Arrow keys for extended selection
   - Ctrl + A to select entire worksheet

2. **Custom Cell Styling**
   ```
   Steps to Change Cell Colors:
   1. Select target cells
   2. Right-click or use Home tab
   3. Select 'Format Cells'
   4. Under 'Fill' tab: Set background to black
   5. Under 'Font' tab: Set text color to white
   ```

### Cell Referencing
1. **Dynamic Workflow**
   - Cells automatically update when referenced cells change
   - Enables automatic calculations
   - Reduces manual data entry errors

2. **Types of References**
   - Relative: Changes when copied (e.g., A1)
   - Absolute: Stays fixed when copied (e.g., $A$1)
   - Mixed: Combination of both (e.g., $A1 or A$1)

## Copy and Paste Operations

### Text to Column Feature
1. **Accessing the Feature**
   ```
   Path: Data Tab â†’ Text to Columns
   ```
   <img width="822" alt="text to column excel" src="https://github.com/user-attachments/assets/3a936bdd-9591-4571-ae8a-0fb2fd342774">


2. **Common Uses**
   - Splitting full names into first and last names
   - Separating addresses into components
   - Converting delimited data into columns

### Paste Special Functions
1. **Basic Options**
   - Values: Paste only the data
   - Formulas: Paste with calculations
   - Formatting: Paste only the appearance

2. **Transpose Function**
   ```
   Steps to Transpose:
   1. Select and copy source data
   2. Right-click destination
   3. Choose 'Paste Special'
   4. Check 'Transpose' box
   ```

## Formula Fundamentals

### Basic Rules
1. **Formula Syntax**
   - Must begin with equals sign (=)
   - Requires mathematical operators:
     * Addition: +
     * Subtraction: -
     * Multiplication: *
     * Division: /

2. **Example Calculations**
   ```excel
   =G4*I4/30    // Multiplies G4 by I4 and divides by 30
   ```

### Formula Visualization
1. **Show Formula Option**
   ```
   Steps to Show Formulas:
   1. Click Formula tab
   2. Select 'Show Formulas'
   - OR -
   Use shortcut: Ctrl + `
   ```

2. **Clear Contents**
   ```
   Steps to Clear:
   1. Select cells
   2. Right-click
   3. Choose 'Clear Contents'
   ```

## Pivot Tables

### Setup Requirements
1. **Data Preparation**
   - Consistent column headers
   - No blank rows or columns
   - Clean, formatted data
   - Date/calendar details in columns

2. **Data Selection Methods**
   ```
   Fast Scrolling Method:
   1. Click first cell
   2. Press Ctrl + Shift + Arrow key
   ```

### Creation Rules
1. **Basic Structure**
   - Highlight complete dataset
   - Insert PivotTable
   - Select destination

2. **Example Configuration**
   ```
   Sample Analysis (Year, Categories, Volumes):
   - Drag Year to Columns area
   - Drag Categories to Rows area
   - Drag Volume to Values area
   ```
<img width="545" alt="Pivot table 1 volume" src="https://github.com/user-attachments/assets/fbc9d67c-1b61-4dba-bb40-f4f63f0152b4">



At a glance you would see the store type with the highest income (convenience store). Its very important to find out why, before drawing conclusions.

<img width="886" alt="pivot table volume with products talent prt" src="https://github.com/user-attachments/assets/12f3bda5-90a6-42ce-b266-83643daa201b">



By adding products to the report, it shows convenience store made more sales, because they had more products in their store.
This is now a good place to start making decisions and asking the right questions.

### Customization and Features
1. **Value Settings**
   - Change summarization from Count to Sum
   - Format numbers with commas
   - Remove decimals for readability

2. **Cross Tabulation**
   - Occurs when totals display in both rows and columns
   - Provides comprehensive data summary
   - Enables multi-dimensional analysis

3. **Data Refresh Requirements**
   ```
   Important: Pivot Tables do not auto-refresh
   
   When to Refresh:
   - After editing source data
   - After adding new records
   - After any modifications to dataset
   
   How to Refresh:
   1. Right-click pivot table â†’ Refresh
   2. Analyze tab â†’ Refresh
   3. Keyboard shortcut: Alt + F5
   ```

## Navigation and Selection

### Keyboard Shortcuts
1. **Movement**
   - Ctrl + Arrow: Move to data edges
   - Ctrl + Home: Go to beginning
   - Ctrl + End: Go to last used cell

2. **Selection**
   ```
   Fast Selection Methods:
   - Ctrl + Shift + Arrow: Select to data edge
   - Ctrl + Space: Select column
   - Shift + Space: Select row
   ```


### Error Prevention
1. **Semantic Errors**
   - Maintain consistent data formats
   - Use appropriate data types
   - Validate formulas before use


## Excel Class - November 13, 2024
---

### Printing a Specific Area
When printing from Excel, the default behavior is to print the entire worksheet, which may not always be desirable. Printing the entire document can lead to wasted paper and ink, especially if you only need to print a specific section of the data. By setting a print area, you can ensure that only the relevant information is printed, making the process more efficient and cost-effective.


**Printing a Specific Screen Area**
-Highlight the area you want to print
-Go to the Page Layout tab
-Click on Print Area and select Set Print Area
-Press Ctrl+P to print the selected area
-After printing, go back to Print Area and select Clear Print Area



### Hyperlinks
Hyperlinks in Excel allow you to create clickable links to various types of content, such as:

1. **Images**: You can link an image to another file, a website, or even an email address. This is useful when you want to quickly access related visual information.

2. **Documents**: Linking a cell or text to another Excel file, Word document, or any other type of file can simplify navigation and access to important resources.

3. **Email Addresses**: Creating a hyperlink to an email address allows users to quickly compose a new message by simply clicking on the link.

4. **Websites**: Linking text or an object to a website URL enables users to quickly access online resources directly from your Excel workbook.

For example, in our Excel class, we might have a worksheet that contains a list of resources. By adding hyperlinks to relevant websites, documents, or email addresses, we can make it easier for our classmates to access this information without having to manually type or search for the links.


**To create a hyperlink:**

-Select the text or object you want to turn into a hyperlink
-Right-click and select Hyperlink...
-Enter the destination (e.g. file path, email address, URL)
-Click OK
-Now when you click the hyperlinked text/object, it will take you to the linked content

### Macros: Automating Repetitive Tasks
In our Excel class, we discussed how macros can help us automate repetitive tasks, such as formatting our worksheets. For instance, let's say we have a standard formatting that we want to apply to all our class worksheets:

1. Set the font to Arial, size 10
2. Remove the gridlines
3. Adjust the column widths to 10

Instead of manually performing these steps every time we create a new worksheet, we can record a macro to automate the process. Here's how we did it:

1. Opened the **Developer** tab
2. Clicked **Record Macro**
3. Named the macro "Formatting" and assigned the shortcut key `Ctrl+M`
4. Performed the formatting steps
5. Clicked **Stop Recording**

Now, whenever we need to apply this standard formatting to a new worksheet, we can simply press `Ctrl+M`, and the macro will automatically execute the formatting steps for us.

Using macros in this way can save us a significant amount of time and effort, allowing us to focus on the more important aspects of our work rather than repetitive, manual tasks.


#  EXCEL Functions (11/18/2024)

## Key Concepts:

1. **Mathematical Formula**: A function is a mathematical formula designed to perform specific calculations in Excel.

2. **Calculations**: Functions automate calculations, allowing us to perform them more efficiently and precisely.

3. **Speed**: Functions enable faster calculations, which is especially useful in a fast-paced work environment.

4. **Common Functions**: Examples include `SUM`, `AVERAGE`, `COUNT`, `COUNTIF`, `SUMIF`, and conditional statements like `IF`.

## Detailed Explanations:

1. **Linear Functions**: These are simple, straightforward functions like `SUM`.

2. **AVERAGE vs. Mean**: While `AVERAGE` calculates the mean, the mean has limitations and can't be fully trusted or relied upon in data analysis.

3. **Skewness**: To address this, introduce skewness using `MODE` and `MEDIAN`. Skewness refers to the assymetry or lack of symmetry in a data distribution. Positive skewness indicates the right tail of the distribution is longer, while negative skewness indicates the left tail is longer. The difference between the mean and median indicates the type of skewness.

4. **Types of Skewness**: There are three main types of skewness:
   - Positive Skewness: The mean is greater than the median, which is greater than the mode. This indicates the distribution has a longer right tail.
   - Negative Skewness: The mean is less than the median, which is less than the mode. This indicates the distribution has a longer left tail. 
   - Zero Skewness: The mean, median, and mode are approximately equal, indicating a symmetric distribution.

5. **Outliers**: Skewness can also indicate the presence of outliers in the data, which are data points that are significantly different from the rest of the distribution.

6. **COUNT vs. COUNTA**: `COUNT` can only count numerical data, while `COUNTA` counts both numerical and non-numerical data.

7. **Conditional Statements**: Functions like `IF` allow the computer to act based on specific criteria (one condition at a time).

8. **Boolean Logic**: Excel uses Boolean logic (true/false) to evaluate these conditional statements.

9. **Parameterized Functions**: All Excel functions are "parameterized," meaning they require input parameters. Understanding this helps you use functions effectively.

10. **Opening Parameters**: To view a function's parameters, simply open the function by typing the function name followed by an open parenthesis. The parameters will then be displayed.

11. **Parameter Range**: Instead of manually typing in cell references, you can select a range of cells using keyboard shortcuts like `Control+Shift+Arrow`.

# Functions Used Today

1. `IF`: Used to categorize individuals as "Adult" or "Teenager" based on their age.
2. `SUMIFS`: Used to perform a sum calculation across multiple criteria.

- `=CONCATENATE(FirstName, " ", LastName)`: Concatenates the first and last name into a full name.
- `=IF([@Age]>25,"Adult","Teenager")`: Checks the age and returns "Adult" or "Teenager" accordingly.
- `=SUMIFS(I4:I15,C4:C15,C15,D4:D15,D15,E4:E15,E15,F4:F15,F15,G4:G15,G15,H4:H15,H15)`: Performs a sum calculation across multiple criteria.

<img width="948" alt="If function to Categorize age" src="https://github.com/user-attachments/assets/557a85f5-ab07-4de1-94e6-22fab83bb311">


=IF([@Age]>25,"Adult","Teenager")
This formula checks the value in the "Age" column for each row. If the age is greater than 25, it returns the text "Adult", otherwise it returns the text "Teenager".
The key components are:

[@Age] - This refers to the "Age" value for the current row.
25 - This is the comparison to check if the age is greater than 25.


"Adult" - The text value to return if the age is greater than 25.
"Teenager" - The text value to return if the age is 25 or less.

This formula can be used to categorize the individuals in the data set based on their age, separating them into "Adult" and "Teenager" groups.



 <img width="952" alt="excel functions" src="https://github.com/user-attachments/assets/907e2f60-2be5-4afb-bf6e-30d8b38e44a4">


1. `MIN`: Returns the smallest numeric value from a range of cells.
2. `MAX`: Returns the largest numeric value from a range of cells.
3. `COUNTA`: Counts the number of cells that contain any data, including numbers, text, and logical values.
4. `COUNTIFS`: Counts the number of cells that meet multiple criteria.
5. `UNIQUE`: Returns a range containing the unique values from a specified range.
6. `COUNTIF`: Counts the number of cells that meet a single criterion.




# Excel Class:Data Preprocessing and Missing Values Analysis
*Date: November 20, 2024*

## Class Theory Notes

### Missing Values Treatment Guidelines
1. Basic Decision Rules:
   - If missing data < 5%: Drop the data
   - If missing data => 5%: Fill with mean value
   - Note: Extrapolation provides more accurate results

**General steps:**
Check all the data set (control A) Find and select blank 
Then from the general search, identify missing values in colmns and do column by column search to fill missing vales.*

### Key Concepts
1. **Converging**
   - Definition: Process of bringing together extrapolated values to get close ranges
   - Purpose: To obtain the closest possible numbers during exploration
   - Application: Used when dealing with multiple extrapolated values
	- Extrapolation: Preferred for more accurate results when filling missing values


2. **Finding Missing Values in Excel**
   - Step 1: Select all cells
     * Use Ctrl + A for precise table area selection
     * Avoid selecting data outside the relevant range
   - Step 2: Navigate to Home â†’ Editing â†’ Find & Select â†’ Blanks â†’ OK
   - Result: All blank cells will be selected

3. **Practical Tips**
   - To highlight missing values:
     * Fill a cell with yellow color
     * Select blanks using Find & Select
     * Use Ctrl + V to paste yellow highlighting
   - To fill missing values:
     * Calculate mean value
     * Select blanks
     * Use Ctrl + V to fill all blanks

   - Quick data movement:
To move data without copy/paste:
     * Hover until four-directional arrow appears
     * Click and drag to move data

## Class Practical Session 


<img width="773" alt="missing values" src="https://github.com/user-attachments/assets/025634c1-ce76-46fa-8df8-09893ce1b20a">

## Missing Values Treatment
### Initial Analysis
- **Focus Area**: Volume column
- **Treatment Method**: Mean imputation, (=Average)
- **Mean Value Used**: 1,435.865
- **Selection Criteria**: Values were missing in specific records

- **Important tip for copying MEAN value to paste on blank celss**
- After calculating mean value, copy before find and select, blank, ok and control V.
- But note copying mean value and pasting will return an error (REF) oR a wrong value, different from the calculated value
- This is because copying the mean from calculated means copying the formular as well, which could change the value.
- To solve this problem, use copy and paste special into another cell (paste special option: values)
- Now values will be pasted, not formular
- This is the one you should copy to perfrom FIND AND SELECT

### Process Steps
1. Data Selection
   - Selected entire dataset using Ctrl+A
   - Focused on Volume column

2. Missing Values Identification
   - Used Find & Select â†’ Blanks
   - Highlighted missing values in yellow for visibility
   
3. Value Imputation
   - Calculated mean value: 1,435.865
   - Applied mean imputation to missing values
   - Verified consistency after imputation
  

## Class: Excel Functions and Data Validation
Date: November 25, 2024

1. Lookup Functions
- VLOOKUP: Vertical column search
- HLOOKUP: Horizontal row search

2. Referencing Techniques
- Absolute Referencing ($A$1): Locks cell reference
- Array Reference: Multiple cell selection
- Range Reference: Continuous cell group

3. Data Validation Overview
A. Definition
- Restricts cell data entry
- Ensures data accuracy
- Prevents invalid inputs

B. Implementation Steps
- Select target cells
- Navigate to Data > Data Validation
- Choose validation criteria
- Set input/error messages
- Define allowed values

C. Validation Types
- Whole Number
- Decimal
- List
- Date Range
- Text Length
- Custom Formula

D. Error Signal Meanings
- Stop (Red X): Blocks entry
- Warning (Yellow !): Cautions user
- Information: Provides guidance

E. Limitations
- Ineffective on:
  * Unlinked cells
  * Protected sheet cells
  * Imported/pasted data

4. Functions Covered today (Practical)
- IF ERROR: Error handling
- ROUNDUP: Rounds up numbers
- Date Functions:
  * DATE
  * YEAR
  * DAY
  * MONTH
  * QUARTILE
- Text Concatenation (&)
- List Validation
- Validation for:
  * Date Ranges
  * Whole Numbers
  * Decimals
  * Text Length


# Advanced Pivot Table/ Creating Models and connecting pipes
## Date: November 27, 2024

In this detailed note, I will cover the key concepts and techniques related to advanced pivot table usage, with a focus on extracting data from columns with different data types and leveraging various functions and features to gain deeper insights.

### 1. Data Types in Columns:
- When working with pivot tables, it's essential to understand the different data types present in the columns, as this affects how you can extract and manipulate the data.
- For numeric data types, you can use functions like `=YEAR()` to extract the year from a date-formatted number (e.g., 201901).
- For text-formatted numeric data, you can use the `=RIGHT(cell, number_of_characters)` function to extract the desired number of characters from the right side of the cell.

### 2. Leveraging Filtering and Search Options:
- Pivot tables offer powerful filtering and search capabilities that allow you to quickly find and analyze specific data points.
- Learn how to utilize the filter options, such as filtering by value, text, or number, to narrow down your analysis and focus on the most relevant information.
- Explore the search function within the pivot table to quickly locate and highlight specific data or insights.

### 3. Handling Data Uniformity:
- Maintaining data uniformity is crucial for accurate analysis and reporting. Ensure that numeric data is consistently formatted, with leading zeros (e.g., 01, 10, 12) instead of varying formats (e.g., 1, 2, 10).

### 4. Understanding Variance and Absolute Variance:
- Pivot tables allow you to calculate both variance and absolute variance, which are essential metrics for understanding the changes in data over time.
- Variance measures the relative change, while absolute variance focuses on the absolute difference between values.
- Analyze these metrics to identify significant shifts in your data and uncover potential areas for improvement or further investigation.

### 5. Interpreting Subtotals and Grand Totals:
- Pivot tables often include subtotals and grand totals, which are calculated columns.
- Subtotals are indicated by a line above the value, while grand totals have a thicker line both above and below the value.
- Understanding these visual cues can help you quickly identify calculated fields and differentiate them from raw data.

### 6. Dynamic Naming and Cross-Sheet Referencing:
- Pivot tables can leverage dynamic naming conventions to make it easier to reference specific data points or calculations.
- Learn how to use cell references to link data from other sheets or workbooks, enabling you to build comprehensive and interconnected financial models.
- 


  <img width="838" alt="FDM analysis" src="https://github.com/user-attachments/assets/4dffc0f3-39dd-42ea-84a1-dba2182253c2">


### 7. Building a Financial Delivery Model (FDM):Class Practise
- A Financial Delivery Model (FDM) was created by selecting key columns and adding data for 2015 and 2016.
- The model includes calculations for net profit, FDM, and gross profit, among other important metrics.
- Analyze the insights from this FDM, focusing on variables, absolutes, and any available slicers to uncover trends, identify areas for improvement, and craft a compelling story-driven report.
- 

  [FDM Report.docx](https://github.com/user-attachments/files/17934208/FDM.Report.docx)



##Important Notes: STEPS
-Create the model (the FDM, add formulars to the models, dont worry about the values. When the pivots are created it will be updated when connected to the pipeline)
_ Create the pipeline (Pivot table)
- Connect the pipeline using dynamic refencing
- Add slicers to the report
- Create the report to tell the story about your model.





