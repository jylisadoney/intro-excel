---
title: Tips
nav: true
---
Please download the <a href="images/IntroExcelSampleWorkbook.xlsx" target="_blank">`IntroExcelSampleWorkbook.xlsx`</a> to follow along with workshop activities.

**Throughout this workshop `Click` refers to a `Left Click` with your mouse**

# Tips and Tricks

Excel offers lots of other ways to organize your spreadsheets and work with your data.

Below are a few tips and tricks that will make use a power-user

## Wrap text
If your data is text-based, you can wrap the text so it appears on multiple lines in a cell.
* Highlight the cell, column, or row where you want to wrap text
* Navigate to the `Home` tab and `Alignment` section
* Click on `Wrap Text`

## Add today's date
* Click on the cell where you want the date to appear
* Hit `Ctrl ; (control key and semicolon)`

## Split text into multiple columns

Let’s say we want to separate our employee’s first names from their last names in the `Salesperson` column.

To do this, navigate to the `Tricks` sheet in our workbook.

Our first step is to insert a new column to the left of `Sales Q1`
* Click on `Column C`
* Right click then click on `Insert`

Then, we can navigate to the `Data` tab and `Data Tools` section
* Click on `Column B` to highlight it
* Click on `Text to Columns`
* Select the `Delimited` opition in the pop-up window
* Click `Next`
* Click to uncheck the box for `Tab`
* Click to check the box for Space
  * We can now see a preview of what our data will look like
* Click `Next`
* Click Finish

Now, since we have two columns with employee names, lets rename `Column B` and `Column C`.
* Click `Cell B1` and rename it as `Employee First Name`
* Click `Cell C1` and rename it as `Employee Last Name`

We can also resize the columns if we need to.
* Click and highlight both `Column B and Column C`
* Navigate to the `Home tab` and `Cells` section
* Click on `Format`
* Click on `Autofit Column Width`