---
title: Advanced
nav: true
---
Please download <a href="images/IntroExcelSampleWorkbook.xlsx" target="_blank">`IntroExcelSampleWorkbook.xlsx`</a> to follow along with workshop activities.

**Throughout this workshop `Click` refers to a `Left Click` with your mouse**

# Advanced Excel Features

Excel offers many other features that you can use to work with your data.

These include `Templates`, `Conditional Formatting`, `Data Validation`, and `Pivot Tables`.

[Data Validation](#data-validation)

# Templates

If you tend to use the same layout in Excel workbooks, you can save it as a `template` and use it again in the future.

## Set the default personal templates location
* Click on `File` then click on `Options`
* Click on the `Save` tab
* Navigate to the `Save workbooks` section
* Type this file path: `C:\Users\[UserName]\Documents\Custom Office Templates`
  * Replace [User Name] with your user name
* Click `OK`

All templates you save to the `My Templates` folder will now appear as options when you create a new Excel workbook

## Save a workbook as a template
* Open Excel
* Open the workbook you want to save as a template
* Click on `File` then click on `Export`
* Under `Export`, click on `Change File Type`
* Under the `Workbook File Types` section, double-click on `Template`
* Type the file name you want to use for the template
* Click on `Save` and close the template

## Create a workbook based on a template
* Open Excel
* Click on `File` then click on `New`
* Click on `Personal`
* Double-click on the template you want to use

Excel will then create a new workbook based on this template.

# Conditional formatting

`Conditional formatting` in Excel lets you format, highlight, or display data that meet certain criteria.

Built in features let you:
* Format cells based on their values
* Format cells that contain certain information
* Format top or bottom values
* Format values that are above or below the average
* Format unique or duplicate values

## Add conditional formatting

Let’s say we want to highlight the `total sales cells` that are `equal or above the average`.
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

We can now see that Excel has highlighted the `total sales cells` that are `equal to or above the average` of our data.

## Edit conditional formatting
* Navigate to the `Home` tab and the `Styles` section
* Click on `Conditional Formatting`
* Click on `Manage Rules`
* Click to highlight the rule you want to edit
* Click on `Edit Rule` and enter your changes

## Remove conditional formatting

### Specific cells
* Click and drag to highlight the `cells, columns, or rows` where you want to remove conditional formatting
* Navigate to the `Home` tab and the `Styles` section
* Click on `Conditional Formatting`
* Click on `Clear Rules` 
* Click on `Clear Rules from Selected Cells`

### Entire sheet
* Navigate to the `Home` tab and the `Styles` section
* Click on `Conditional Formatting`
* Click on `Clear Rules` 
* Click on `Clear Rules from Entire Sheet`

# Data validation

Data validation in Excel is used to restrict what can be entered into a cell.

* Click to highlight the cell(s) where you want to add validation
* Navigate to the `Data` tab and the `Data Tools` section
* Hover and click the `Data Validation` icon

These first three steps work for any data validation.

## Restrict data entry to whole numbers within limits
* In the `Settings` tab:
  * Set the `Allow` drop-down box to `Whole number`
  * Set the `Data` drop-down box to `between`
  * Enter the `Minimum` and `Maximum` range of whole numbers to allow
    * For example, Minimum: 1 and Maximum: 25
* Click `OK`

Now, if you try to enter a number **outside the specified range**, you will see an error message.

You can even specify the `Input Message` or `Error Alert` if you choose.

## Restrict data entry to a range of dates
* In the `Settings` tab:
  * Set the `Allow` drop-down box to `Date`
  * Set the `Data` drop-down box to `between`
  * Enter the `Start date` and `End date` to allow
    * For example, Start date: 01/01/1999 and End date: =TODAY()
* Click `OK`

Now, if you try to enter a date that occurred **before** 01/01/1999 or **after** today, you will see an error message.

`=TODAY()` will automatically update based on the actual date.

## Restrict data entry to a drop-down list of data values
* Click to highlight the cell or cells you want to validate
* Navigate to the `Data` tab and the `Data Tools` section
* Hover and click the `Data Validation` icon
* In the `Settings` tab:
  * Set the `Allow` drop-down box to `List`
  * In the `Source` box, type your data values and separate them with commas without spaces
    * For example, `Yes,No,Maybe`
* Click to check the `In-cell dropdown` box
* Click `OK`

Now, when you want to enter text into that cell(s), you can select from the **allowed options**.

## Remove data validation 	
* Click to highlight the cell(s) where you want to remove validation
* Navigate to the `Data` tab and the `Data Tools` section
* Hover and click the `Data Validation` icon
* Click `Clear All`
* Click `OK`

# Pivot tables

# More Help
## Conditional formatting
Excel includes numerous ways to use `conditional formatting`. To learn more about them, navigate to the `Home` tab and the `Styles` section. 

You can also visit <a href="https://support.office.com/en-us/article/Enter-and-format-data-fef13169-0a84-4b92-a5ab-d856b0d7c1f7#ID0EAABAAA=Conditional_formatting" target="_blank">Microsoft's Enter and Format Data website</a>.

## Data validation
Visit <a href="https://support.office.com/en-us/article/apply-data-validation-to-cells-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249" target="_blank">Microsoft's Apply Data Validation to Cells</a>.