---
title: Advanced
nav: true
---
Please download <a href="images/IntroExcelSampleWorkbook.xlsx" target="_blank">`IntroExcelSampleWorkbook.xlsx`</a> to follow along with workshop activities.

**Throughout this workshop `Click` refers to a `Left Click` with your mouse**

# Advanced Excel Features

Excel offers many other features that you can use to work with your data.

These include [Templates](#templates) and [Data Validation](#data-validation). Links to [more help](#more-help) on each of these features, including [PivotTables](#pivottables), is available at the bottom of the page.

# Templates

If you often use the same workbook layout, you can save it as a `template` and use it again in the future.

To use templates, we first have to `set the default personal templates location`.

## Set the default personal templates location
* Click on `File` then click on `Options`
* Click on the `Save` tab
* Navigate to the `Save workbooks` section
* Type this file path: `C:\Users\[UserName]\Documents\Custom Office Templates`
  * Replace [User Name] with your user name
* Click `OK`

Unless you want to change the `default personal templates location`, you will only need to set this once.

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

Excel will then create a new workbook based on this template. Be sure to `Save` this new workbook.

# Data validation

Data validation in Excel controls `what` you can enter into a cell.

These first three steps are the same for any type of `data validation`.
* Click and drag to highlight the cell(s) where you want to add data validation
* Navigate to the `Data` tab and the `Data Tools` section
* Hover and click the `Data Validation` icon

## Restrict data entry to whole numbers within limits
* Click and drag to highlight the cell(s) where you want to add data validation
* Navigate to the `Data` tab and the `Data Tools` section
* Hover and click the `Data Validation` icon
* In the `Settings` tab:
  * Set the `Allow` drop-down box to `Whole number`
  * Set the `Data` drop-down box to `between`
  * Enter the allowed `Minimum` and `Maximum` range of whole numbers
    * For example, Minimum: 1 and Maximum: 25
* Click `OK`

Now, if you try to enter a number **outside the specified range**, you will see an error message.

You can even specify the `Input Message` or `Error Alert` if you choose.

## Restrict data entry to a range of dates
* Click and drag to highlight the cell(s) where you want to add data validation
* Navigate to the `Data` tab and the `Data Tools` section
* Hover and click the `Data Validation` icon
* In the `Settings` tab:
  * Set the `Allow` drop-down box to `Date`
  * Set the `Data` drop-down box to `between`
  * Enter the allowed `Start date` and `End date`
    * For example, Start date: 01/01/1999 and End date: =TODAY()
* Click `OK`

Now, if you try to enter a date **before** 01/01/1999 or **after** today, you will see an error message.

`=TODAY()` will automatically update based on the actual date.

## Restrict data entry to a drop-down list of data values
* Click and drag to highlight the cell(s) where you want to add data validation
* Navigate to the `Data` tab and the `Data Tools` section
* Hover and click the `Data Validation` icon
* Click to highlight the cell or cells you want to validate
* Navigate to the `Data` tab and the `Data Tools` section
* Hover and click the `Data Validation` icon
* In the `Settings` tab:
  * Set the `Allow` drop-down box to `List`
  * In the `Source` box, type your data values and **separate them with commas without spaces**
    * For example, `Yes,No,Maybe`
* Click to check the `In-cell dropdown` box
* Click `OK`

Now, when you want to enter text, you can only select from the **allowed options**.

## Remove data validation 	
*Click and drag to highlight the cell(s) where you want to remove data validation
* Navigate to the `Data` tab and the `Data Tools` section
* Hover and click the `Data Validation` icon
* Click `Clear All`
* Click `OK`

# More Help
## Templates
To learn more about `templates`, visit <a href="https://support.office.com/en-us/article/save-a-workbook-as-a-template-58c6625a-2c0b-4446-9689-ad8baec39e1e" target="_blank">Microsoft's Save a Workbook as a Template website</a>.

## Data validation
To learn more about `data validation`, navigate to the `Data` tab and the `Data Tools` section. 

You can also visit <a href="https://support.office.com/en-us/article/apply-data-validation-to-cells-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249" target="_blank">Microsoft's Apply Data Validation to Cells website</a>.

## PivotTables
PivotTables offer greater control and flexibility when presenting and analyzing data. To learn more about `PivotTables`, navigate to the `Insert` tab and the `Charts` section.

You can also visit <a href="https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables" target="_blank">Microsoft's PivotTables website</a>.
