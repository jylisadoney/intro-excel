---
title: Tips and Tricks
nav: true
---

# Tips and Tricks

Excel offers lots of other ways to organize your spreadsheets and work with your data.

Below are a few tips and tricks that will make use a power-user

## Add today's date
* Left click on the cell where you want the date to appear
* Hit `Ctrl ; (control key and semicolon)`

## Split text into multiple columns

Let’s say we want to separate our employee’s first names from their last names in the `Salesperson` column.

To do this, navigate to the `Tricks` sheet in our workbook.

Our first step is to insert a new column to the left of `Sales Q1`
* Left click on `Column C`
* Right click then and left click on `Insert`

Then, we can navigate to the `Data` tab and `Data Tools` section
* Left click on `Column B` to highlight it
* Left click on `Text to Columns`
* Select `Delimited` in the pop-up window
* Click `Next`
* Uncheck the box for `Tab`
* Check the box for Space
  * We can now see a preview of what our data will look like
* Click `Next`
* Click Finish

Now, since we have two columns with employee names, lets rename `Column B` and `Column C`.
* Left click `Cell B1` and rename it as `Employee First Name`
* Left click `Cell C1` and rename it as `Employee Last Name`

We can also resize the columns if we need to.
* Left click and highlight both `Column B and Column C`
* Navigate to the `Home tab` and `Cells` section
* Left click on `Format`
* Left click on `Autofit Column Width`

## Additional functions

### Math & Trig
* Difference (no function exists)
  * Example: =B1-B2
* Divide (no function exists)
  * Example: =B1/B2 
* [=PRODUCT](https://support.office.com/en-us/article/product-function-8e6b5b24-90ee-4650-aeec-80982a0512ce): multiplies cells
* [=SUM](https://support.office.com/en-us/article/sum-function-043e1c7d-7726-4e80-8f32-07b23e057f89) adds cells

### Statistical
* [=AVERAGE](https://support.office.com/en-us/article/average-function-047bac88-d466-426c-a32b-8f33eb960cf6): identify the average/mean for a range of cells
* [=COUNTIF](https://support.office.com/en-us/article/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34): counts the number of cells in a range that meet a specific criterion 
* [=MAX](https://support.office.com/en-us/article/max-function-e0012414-9ac8-4b34-9a47-73e662c08098): identify the largest number contained in a range of cells
* [=MEDIAN](https://support.office.com/en-us/article/median-function-d0916313-4753-414c-8537-ce85bdd967d2): identify the median for a range of cells
* [=MIN](https://support.office.com/en-us/article/min-function-61635d12-920f-4ce2-a70f-96f202dcc152): identify the smallest number contained in a range of cells
* [=MODE](https://support.office.com/en-us/article/mode-function-e45192ce-9122-4980-82ed-4bdc34973120): identify the mode for a range of cells
* [=STDEV](https://support.office.com/en-us/article/stdev-function-51fecaaa-231e-4bbb-9230-33650a72c9b0): identify the standard deviation for a range of cells

### Text
* [=CONCATENATE](https://support.office.com/en-us/article/concatenate-function-8f8ae884-2ca8-4f7a-b093-75d702bea31d): merges data from two cells into one cell
* [=LOWER](https://support.office.com/en-us/article/lower-function-3f21df02-a80c-44b2-afaf-81358f9fdeb4): converts text to lowercase
* [=PROPER](https://support.office.com/en-us/article/proper-function-52a5a283-e8b2-49be-8506-b2887b889f94): capitalizes the first letter in each word
* [=UPPER](https://support.office.com/en-us/article/upper-function-c11f29b3-d1a3-4537-8df6-04d0049963d6): converts text to uppercase
