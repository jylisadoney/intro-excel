---
title: Advanced
nav: true
---
Please download the <a href="images/IntroExcelSampleWorkbook.xlsx" target="_blank">`IntroExcelSampleWorkbook.xlsx`</a> to follow along with workshop activities.

# Advanced Excel Features

Excel offers many other features that you can use to work with your data.

These include `Templates`, `Conditional Formatting`, `Data Validation`, and `Pivot Tables`.

# Templates

If you tend to use the same layout in Excel workbooks, you can save it as a `template` and use it again in the future.

## Set the default personal templates location
* Left click on `File` then left click on `Options`
* Left click on the `Save` tab
* Navigate to the `Save workbooks` section
* Type this file path: `C:\Users\[UserName]\Documents\Custom Office Templates`
  * Replace [User Name] with your user name
* Left click `OK`

All templates you save to the `My Templates` folder will now appear as options when you create a new Excel workbook

## Save a workbook as a template
* Open the workbook you want to save as a template
* Left click on `File` then left click on `Export`
* Under `Export`, left click on `Change File Type`
* Under the `Workbook File Types` section, double-click on `Template`
* Type the file name you want to use for the template
* Left click on `Save` and close the template

## Create a workbook based on a template
* Open Excel
* Left click on `File` then left click on`New`
* Left click on `Personal`
* Double-click on the template you want to use

Excel will then create a new workbook based on this template.

# Conditional formatting

Conditional formatting in Excel lets you format, highlight, or display data in your workbook that meets certain criteria.

Built in features let you:
* Format cells based on their values
* Format cells that contain certain information
* Format top or bottom values
* Format values that are above or below the average
* Format unique or duplicate values

## Add conditional formatting

Let’s say we want to highlight the total sales cells that are equal or above the average.

To do this, navigate to the `CondFormat` sheet in our workbook.
* Navigate to the `Home` tab and the `Styles` section
* Left click on `Column E` to highlight the Total Sales data
* Left click on `Conditional Formatting`
* Left click on `New Rule`
* Left click on `Format only values that are above or below average`
* In the drop-down box, select `equal or above`

Now we have to set what we want the formatting to look like.
* Left click on `Format`
* Left click on `Fill`
* Left click the color you want to use
* Left click `OK`, then click `OK` again

We can now see that Excel has highlighted the cells that include `Total Sales` that are equal to or above the average of our data.

## Edit conditional formatting
* Navigate to the `Home` tab and the `Styles` section
* Left click on `Conditional Formatting`
* Left click on `Manage Rules`
* Left click to highlight the rule you want to edit
* Left click on `Edit Rule` and enter your changes

## Remove conditional formatting

### Specific cells
* Navigate to the `Home` tab and the `Styles` section
* Left click and drag to highlight the `cells, columns, or rows` where you want to remove conditional formatting
* Left click on `Conditional Formatting`
* Left click on `Clear Rules` 
* Left click on `Clear Rules from Selected Cells`

### Entire sheet
* Navigate to the `Home` tab and the `Styles` section
* Left click on `Conditional Formatting`
* Left click on `Clear Rules` 
* Left click on `Clear Rules from Entire Sheet`

# Data validation


# Pivot tables

# More Help
## Conditional formatting
Excel includes numerous ways to use `conditional formatting`. To learn more about them, navigate to the `Home` tab and the `Styles` section. 

You can also visit <a href="https://support.office.com/en-us/article/Enter-and-format-data-fef13169-0a84-4b92-a5ab-d856b0d7c1f7#ID0EAABAAA=Conditional_formatting" target="_blank">Microsoft's Enter and Format Data website</a>.