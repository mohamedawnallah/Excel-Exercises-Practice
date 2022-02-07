# **Excel Skills for Business**

## _Excel Skills for Business Essentials_:

## 1st Week:

- Excel User Interface
- Key Terminology
- Selecting and Navigation
- Working with Data
- Fill Handle

## 2nd Week:

- Formulas
- Functions
- More Functions
- Relative &amp; Absolute Cell References

## 3rd Week:

- Font Formatting
- Borders
- Alignment
- Format Painter
- Number Formatting
- Styles &amp; Themes

## 4th Week:

- Managing rows &amp; columns
- Find &amp; replace
- Filter &amp; sort
- Conditional Formatting

## 5th Week:

- Print preview
- Margins, scale &amp; orientation
- Page breaks
- Headers &amp; Footers

## 6th Week:

- Basic Charts
- Chart Styles
- Chart Types

-----

## _Excel Skills for Business_: Intermediate I:

## 1st Week:

- Multiple Worksheets:
  - Group Sheets
  - Sheets, tabs manipulations (Deleting - Inserting - Copying - moving -renaming - hiding - unhiding)
- 3D Formulas:
  - 3D Formulas vs 2D Formulas
  - Summing values across multiple worksheets where the structure is the same
- Linking Workbooks:
  - Arrange all windows so you can see them simultaneously
  - Linking multiple workbooks
  - The link consists of :
    - Workbook name + Sheet name + absolute reference of cell required to link
- Consolidation Values by position and by reference:
  - Not updated dynamically of data
  - You can run it multiple times to update data.
  - Consolidation by References (Categories) from different worksheets that may not be typically identical

## 2nd Week:

- Combining Text:
  - Concat Function
  - Ampersand symbol
- Text Case:
  - LOWER Function
  - UPPER Function
  - PROPER Function
- Extracting Text:
  - LEFT Function
  - MID Function
  - Right Function
  - MID(text,start\_num,num\_chars)
- Finding Text:
  - FIND(find\_text,within\_text,num\_chars)
- Date Calculations:
  - NOW Function
  - TODAY Function
  - YEARFRAC Function
  - **CTRL + ;** to insert current date in a cell
  - **CTRL + :** to insert current time in a cell

## 3rd Week:

- Named Ranges:
  - Name Box (Don&#39;t Use Hyphen or a space)
  - Any three letters followed by a number gives us a cell reference
  - Named Range gives us automatically the absolute cell reference
  - Name Box is not used to change the named range
- Create and Manage Ranges:
  - CMD + SHIFT + Down Arrow to select data in column till you hit a blank cell
  - Three ways to create name ranges:
    - Name Box
    - Define Names from Formula Tab
    - Create from Selection from Formula Tab
  - Managing Named Ranges:
    - Name Manager in Formulas Tab (Not for Mac Users) but Mac users can get simpler version through Define names in formula tab
- Apply Ranges to Formulas:
  - To get all named ranges, press F3
- Apply Names:
  - CMD + SHIFT + F3 -\&gt; Create Named Ranges from Selection
  - Apply Names is sublabeled in Define Name button in Formulas tab
  - As our workbooks gets complex it becomes more necessary to past all named ranges list by clicking F3 -\&gt; Paste List
  - Implicit Difference &amp; At Symbol @

## 4th Week:

- COUNT and COUNTIFS:
  - To extract some meaning from the data, we need to create these summary reports
  - COUNTIFS count the number of times &quot;certain criterias&quot; is met
  - SUMIFS sums a range based on &quot;certain criterias&quot;
  - In absolute cell references, we lock the column and we lock the row
  - Mixed cell references, we may lock the column but not the row or we can lock the row, but not the column
  - COUNT function counts all cells that contain a numeric value all empty cells are ignored
  - COUNTA Function counts all cells that contain both numerical and alphanumerical data all empty cells are ignored
  - COUNTBLANK function counts all empty cells
  - COUNTIFS needs two arguments (at least):
    - Criteria\_Range
    - Criteria1
    - For COUNTIFS, text and expressions must be within quotation marks
- SUMIFS(More than one criteria ranges, criteria):
  - SUM\_RANGE
  - CRITERIA\_RANGE 1
  - CRITERIA1
- Sparklines:
  - Sparklines are min-charts that fit in a Single Cell
  - Insert Tab then Sparklines
  - Sparklines are grouped together
- Advanced Charting:
  - Switching Row / Column
  - Select Data
  - Change Chart Type
  - Secondary Axis
  - Combo Chart (supported with line one overlay first axis)
- Trendlines:
  - Adjusting Vertical Scale through Formatting Axis
  - Add Trendline for Exponential, Linear, Logarithmic, Power Functions
  - R^2 closer to one the more it&#39;s linear

## 5th Week:

- Create and Format Tables:
  - Table:
    - Hold database-style data
    - Column represents a field that contains a particular type of information
    - Record represents entire information about customer, or one transaction
    - Table Tab:
      - Contextual Ribbon Tab
      - Table Name
      - Banded Rows
      - …etc features
    - Tables vs Ranges:
      - Tables:
        - Selection of Data is easier a lot than in Ranges
        - No need to worry about Freezing Panes when working with Tables
        - We&#39;re really looking to use them whenever we&#39;re working with raw data entries
        - It comes to trouble, when you&#39;re looking to create summaries of sections inside your table
      - Ranges:
        - Selection of Data is harder than in Tables
        - You need to worry about Freezing panes
        - Subtotal tool is available here
- Sort and Filter in Tables:
  - Removing Duplicates:
    - Table =\&gt; Remove Duplicates
  - Accessing Sort Function:
    - Sorting:
      - Table =\&gt; Filter Button
      - Data Tab
      - Home Tab
      - Right Click =\&gt; Sort
      - Header Row
  - Total Row
  - It is good practice to always clear filters,especially when you work in a shared Environment
- Automation:
  - Named Ranges are Automated with Table
  - CTRL SHIFT + Shortcut for adding new Row
  - Named Ranges Extend Automatically in Tables Not in Ranges
  - Charts Update automatically when they are attached to tables Unlike Ranges
- Subtotalling:
  - Feature in Excel that summarizes subsections of your data
  - We need to convert our table back to range
  - Subtotal creates outline on the left