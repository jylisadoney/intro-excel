---
title: Insert
nav: true
---
Please download the <a href="images/IntroExcelSampleWorkbook.xlsx" target="_blank">`IntroExcelSampleWorkbook.xlsx`</a> to follow along with workshop activities.

# Insert Functions

In Excel, functions can help you do math, calculate standard deviation, edit text, or find specific types of data.

Today, we are going to focus on two functions: `SUM` and `PROPER`.

## SUM function

The SUM function (a Math & Trig function), adds values together. 

Let’s say we want to see the total amount of sales made by employees in Q1 and Q2.

To do this, we first have to insert a new column after `Sales Q2 (Column D)`.
* Left click on `Column E` to highlight it
* Right click then left click on `Insert` to add a column
* Type `Total Sales` in `Cell E1` and resize the column if necessary

Now we can calculate the total amount of sales with the `SUM` function.
* Type `=SUM(C2,D2)`
  * Option 1: Type `C2,D2` **OR**
  * Option 2: Left click on `C2`, type a `comma (,)` then left click on `D2`
* Hit `Enter` on your keyboard to `SUM` these two cells together

To easily sum the remaining employee sales, we can use `Autofill`.
* Left click on `Cell E2`
* Hover over the bottom right corner of the cell until you get a **black plus sign (+)**
* Left click when you see the `black plus sign (+)`
* Drag to highlight the entire column to `Cell E10`, then unclick

## PROPER function

As we are looking at our data, we can see that the state names in `Column A` are all lowercase.

To make our spreadsheet more visually appealing, we can use the PROPER function (a Text function) to change this text to proper capitalization. This will capitalize the first letter of the word and convert all other letters to lowercase.

To do this, we first have to insert a new column before `Salesperson (Column B)`
* Left click on `Column B` to highlight it
* Right click then left click on `Insert` to add a column

Now, we can change the capitalization with the `PROPER` function
* Left click on `Cell B1`
* Type =proper(A1)
  * Option 1: Type `A1` **OR**
  * Option 2: Left click on `A1`
* Hit `Enter` on your keyboard

To easily capitalize the remaining state names, we can use `Autofill`.
* Left click on `Cell B1`
* Hover over the bottom right corner of the cell until you get a *black plus sign (+)*
* Left click when you see the `black plus sign (+)`
* Drag to highlight the entire column to `Cell B10`, then unclick

Our next step is to place these state names back in `Column A` using a special `copy and paste`.
* Left click on `Column B` to highlight it
* Right click on `Column B`
* Left click on `Copy`
* Left click on `Column A` to highlight it
* Right click on `Column A` 
* Left click on `Paste Special`
* Under `Paste`, left click on `Values`
  * This will let us copy and paste just the text

Since `Column B` is now duplicated, we can delete it.
* Left click on `Column B` to highlight it
* Right click then left click on `Delete`

# More Help
Excel includes numerous other functions. To learn more about them, navigate to the `Formulas` tab and the `Function Library` section. 

You can also visit <a href=" https://support.office.com/en-us/article/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb?ui=en-US&rs=en-US&ad=US" target="_blank">Microsoft's Excel Functions (by Category) website</a>.