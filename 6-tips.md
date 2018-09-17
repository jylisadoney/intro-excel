---
title: Tips
nav: true
---
Please download <a href="images/IntroExcelSampleWorkbook.xlsx" target="_blank">`IntroExcelSampleWorkbook.xlsx`</a> to follow along with workshop activities.

**Throughout this workshop `Click` refers to a `Left Click` with your mouse**

# Tips and Tricks

Excel offers many other ways to organize your sheets and work with your data.

We will discuss how to [split text into multiple columns](#split-text-into-multiple-columns), [wrap text](#wrap-text), and [add the date and time](#add-the-date-and-time).

## [Split text into multiple columns](#split-text-into-multiple-columns)

Let’s say we want to separate our employee’s first and last names in the `Salesperson` column.

To do this, navigate to the `Employee Sales sheet (or Sheet1)` in our workbook.

Our first step is to insert a new column to the left of `Sales Q1 (Column C)`.
* Click on `Column C`
* Right click then click on `Insert`

Then, navigate to the `Data` tab and `Data Tools` section
* Click on `Column B` to highlight it
* Click on `Text to Columns`
* Select the `Delimited` option in the pop-up window
* Click `Next`
* Click `Tab` to uncheck the box
* Click to check the box for `Space`
  * We can now see a preview of what our data will look like
* Click `Next`
* Click `Finish`

If you see a `There's already data here. Do you want to replace it?` pop-up, you will need to insert additional blank columns to ensure that the `split text` does not replace other data.

Now, since we have two columns with employee names, let's rename `Column B` and `Column C`.
* Click `Cell B1` and rename it to `Employee First Name`
* Click `Cell C1` and rename it to `Employee Last Name`

We can also `resize the columns` if we need to.
* Click and highlight both `Column B and Column C`
* Navigate to the `Home tab` and `Cells` section
* Click on `Format`
* Click on `Autofit Column Width`

## [Wrap text](#wrap-text)
If your data is text-based, you can `wrap the text` so it appears on multiple lines in a cell.
* Click to highlight the cell, column, or row where you want to wrap text
* Navigate to the `Home` tab and `Alignment` section
* Click on `Wrap Text`

## [Add the date and time](#add-the-date-and-time)
### Add today's date (static)
* Click on the cell where you want the date to appear
* Hit `Ctrl ;` on your keyboard (same for Mac users)
* Hit `Enter` on your keyboard

### Add the current time (static)
* Click on the cell where you want the time to appear
* Hit `Ctrl Shift ;` on your keyboard (`Command ;` for Mac users)
* Hit `Enter` on your keyboard

### Add today's date and current time (static)]
* Click on the cell where you want the time to appear
* Hit `Ctrl ;` on your keyboard (same for Mac users)
* Hit the `Space bar` on your keyboard (same for Mac users)
* Hit `Ctrl Shift ;` on your keyboard (`Command ;` for Mac users)
* Hit `Enter` on your keyboard

### Add a date or time that updates (dynamic)
* Click on the cell where you today's date and current time to appear
* Type: `=NOW()`
* Hit `Enter` on your keyboard

Now the `current date and time` will update when the workbook/sheet is opened.
