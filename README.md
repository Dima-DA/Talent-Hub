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
   - Insert â†’ PivotTable
   - Select destination

2. **Example Configuration**
   ```
   Sample Analysis (Year, Categories, Volumes):
   - Drag Year to Columns area
   - Drag Categories to Rows area
   - Drag Volume to Values area
   ```

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

## Data Management

### Error Prevention
1. **Semantic Errors**
   - Maintain consistent data formats
   - Use appropriate data types
   - Validate formulas before use

   ```



