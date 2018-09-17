---
title: Insert
nav: true
---
Please download <a href="images/IntroExcelSampleWorkbook.xlsx" target="_blank">`IntroExcelSampleWorkbook.xlsx`</a> to follow along with workshop activities.

**Throughout this workshop `Click` refers to a `Left Click` with your mouse**

# Insert Functions

Functions can help you do math, calculate standard deviation, edit text, or find specific types of data.

Today, we are going to focus on two functions: `[SUM](#sum)` and `[PROPER](#proper)`.

To start, let's navigate to the `Employee Sales` sheet (or Sheet1).

## [SUM function](#sum)

The `SUM` function (a Math & Trig function) adds values together. 

Let’s say we want to see the `total amount of sales` made by employees in Q1 and Q2.

To do this, we first have to insert a new column to the left of `Hire Date (Column E)`.
* Click on `Column E` to highlight it
* Right click then click on `Insert` to add a column
* Type `Total Sales` in `Cell E1` and resize the column if necessary
  * You can use the same steps to insert rows.

Now we can use the `SUM` function to add `Sales Q1` and `Sales Q2` together.
* Click on `Cell E2` 
* Type `=SUM(C2,D2)`
  * Option 1: Type `C2,D2` **OR**
  * Option 2: Click on `C2`, type a `comma (,)` then click on `D2`
* Hit `Enter` on your keyboard to `SUM` these two cells together

To easily sum the remaining employee sales, we can use `Autofill`.
* Click on `Cell E2`
* Hover over the bottom right corner of the cell until you get a **black plus sign (+)**
* Click and hold when you see the `black plus sign (+)`
* Drag the `fill handle` to `Cell E10`, then unclick

## [PROPER function](#proper)

As we are looking at our data, we can see that the state names in `Column A` are all lowercase.

To make our spreadsheet more visually appealing, we can use the `PROPER` function (a Text function) to change this text to proper capitalization. This will capitalize the first letter of each word and convert all other letters to lowercase.

To do this, we first have to insert a new column before `Salesperson (Column B)`
* Click on `Column B` to highlight it
* Right click then click on `Insert` to add a column

Now, we can change the capitalization with the `PROPER` function
* Click on `Cell B1`
* Type =PROPER(A1)
  * Option 1: Type `A1` **OR**
  * Option 2: Click on `A1`
* Hit `Enter` on your keyboard

To easily capitalize the remaining state names, we can use `Autofill`.
* Click on `Cell B1`
* Hover over the bottom right corner of the cell until you get a **black plus sign (+)**
* Click and hold when you see the `black plus sign (+)`
* Drag the `fill handle` to `Cell B10`, then unclick

Our next step is to place these state names back in `Column A` using a special `copy and paste`.
* Click on `Column B` to highlight it
* Right click on `Column B`
* Click on `Copy`
* Click on `Column A` to highlight it
* Right click on `Column A` 
* Click on `Paste Special`
* Under `Paste`, click on `Values`
  * This will let us copy and paste just the text

Since `Column B` is now duplicated, we can delete it.
* Click on `Column B` to highlight it
* Right click then click on `Delete`

# More Help
To learn more about `functions`, navigate to the `Formulas` tab and the `Function Library` section. 

You can also visit <a href=" https://support.office.com/en-us/article/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb?ui=en-US&rs=en-US&ad=US" target="_blank">Microsoft's Excel Functions (by Category) website</a>.
