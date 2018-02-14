---
title: Tips and Tricks
nav: true
---
Please download the <a href="images/IntroExcelSampleWorkbook.xlsx" target="_blank">`IntroExcelSampleWorkbook.xlsx`</a> to follow along with workshop activities.

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