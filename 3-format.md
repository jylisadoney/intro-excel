---
title: Format
nav: true
---
Please download <a href="images/IntroExcelSampleWorkbook.xlsx" target="_blank">`IntroExcelSampleWorkbook.xlsx`</a> to follow along with workshop activities.

**Throughout this workshop `Click` refers to a `Left Click` with your mouse**

# Format and Sort Data

After we've organized our spreadsheet, we can work on formatting and sorting our data. We will discuss how to [format dates](#format-dates), [format currency](#format-currency), [sort data](#sort-data), and add [conditional formatting](#conditional-formatting).

# [Format dates](#format-dates)
In Excel, dates are easier to organize and sort when they are represented numerically.

Let's fix our `Hire Date` column to make it more readable.

* Navigate to the `Employee Sales` sheet (or Sheet1)
* Click on `Column E` to highlight it
* Navigate to the `Home` tab and the `Number` section
* Click on the arrow in the bottom right of the `Number` section
* Under Category, click on `Date`
* Scroll down and click on `03/14/2012`
* Click `OK`

The dates in the `Hire Date` columns are now represented by numbers.

# [Format currency](#format-currency)
We can also format our `Sales Q1` and `Sales Q2` columns so that the data are displayed as Currency.
* Click on `Column C`
* Drag to highlight `Column D` (both columns should now be highlighted)
* Navigate to the `Home` tab and the `Number` section
* Click on the Drop-Down box
* Click on `Currency`

The data in our `Sales Q1` and `Sales Q2` columns are now formatted as currency.

# [Sort data](#sort-data)
Now let’s say we want to organize our sheet based on `when an employee joined the company`.
* Highlight all of the data
  * Option 1: Click on `Cell A1` and drag to highlight all data **OR**
  * Option 2: Hit `Ctrl A` on your keyboard (`Command A` for Mac users)
* Navigate to the `Data` tab and the `Sort & Filter` section
* Slick on `Sort`
* Check the `My data has headers` box
* In the `Sort by` drop-down box, click on `Hire Date`
  * Set the `Sort on` drop-down box to `Values`
  * Set the `Order` drop-down box to `Oldest to Newest`
* Click on `OK`

Now the data in our sheet are sorted based on Hire Date. We could also sort `Hire Date` as `Newest to Oldest`.

# [Conditional Formatting](#conditional-formatting)

`Conditional formatting` in Excel lets you:
* Format cells based on their values
* Format cells that contain certain information
* Format top or bottom values
* Format values that are above or below the average
* Format unique or duplicate values

## Add conditional formatting

Let’s say we want to highlight the `Total Sales cells` that are `equal or above the average`.
* Navigate to the `CondFormat` sheet in our workbook.
* Click on `Column E` to highlight the `Total Sales data`
* Navigate to the `Home` tab and the `Styles` section
* Click on `Conditional Formatting`
* Click on `New Rule`
* Click on `Format only values that are above or below average`
* In the drop-down box, select `equal or above`

Now we have to set what we want the formatting to look like.
* Click on `Format`
* Click on `Fill`
* Click the color you want to use
* Click `OK`, then click `OK` again

We can now see that Excel has highlighted the `Total Sales cells` that are `equal or above the average` of our data.

## Edit conditional formatting
* Navigate to the `Home` tab and the `Styles` section
* Click on `Conditional Formatting`
* Click on `Manage Rules`
* Click to highlight the rule you want to edit
* Click on `Edit Rule` and enter your changes

## Remove conditional formatting

### Specific cells
* Click and drag to highlight the cells(s) where you want to remove conditional formatting
* Navigate to the `Home` tab and the `Styles` section
* Click on `Conditional Formatting`
* Click on `Clear Rules` 
* Click on `Clear Rules from Selected Cells`

### Entire sheet
* Navigate to the `Home` tab and the `Styles` section
* Click on `Conditional Formatting`
* Click on `Clear Rules` 
* Click on `Clear Rules from Entire Sheet`

# More Help

## Format data
To learn more about `formatting data`, navigate to the `Home` tab and the `Numbers` section. 

You can also visit <a href="https://support.office.com/en-us/article/enter-and-format-data-fef13169-0a84-4b92-a5ab-d856b0d7c1f7?ui=en-US&rs=en-US&ad=US#ID0EAABAAA=Format_data" target="_blank">Microsoft's Enter and Format Data website</a>.

## Sort data
To learn more about `sorting data`, navigate to the `Data` tab and the `Sort & Filter` section.

You can also visit <a href=" https://support.office.com/en-us/article/import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde?ui=en-US&rs=en-US&ad=US#ID0EAABAAA=Sort_and_filter" target="_blank">Microsoft's Import and Analyze Data website</a>.

## Add conditional formatting
To learn more about `conditional formatting`, navigate to the `Home` tab and the `Styles` section. 

You can also visit <a href="https://support.office.com/en-us/article/Enter-and-format-data-fef13169-0a84-4b92-a5ab-d856b0d7c1f7#ID0EAABAAA=Conditional_formatting" target="_blank">Microsoft's Enter and Format Data website</a>.
