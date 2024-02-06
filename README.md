# Excel-Tips-and-Tricks
This a compilation of the most commonly used functions, VBA, and Power Query tricks I use frequently to make data analysis just a little bit easier!

## Content
* [Arithmetic Functions](#Arithmetic-Functions)
  * [Common Arithmetic Functions](#Common-Arithmetic-Functions)
      * [Using A Function To Solve A Problem](#Find-the-subtotal-SUM-amount-of-Money-Spent-Column-D-using-the-SUBTOTAL-function)
* [Logical Operation Functions](#Logical-Operation-Functions)
  * [Common Operation Functions](#Common-Operation-Functions)
      * [Merged Function Example](#Example-of-a-merged-logical-operation-function) 
      * [Common "IF()" Variations](#Common-IF-Variations)
      * [Using A Function To Solve A Problem](#check-the-average-value-of-money-spent-column-d-by-france-column--row-a8-using-the-averageif-function)
* [Logical Operation Functions "IFS()"](#logical-operation-functions-IFS)
  * [Common Operation Functions "IFS()"](#Common-IFS-Variations) 
      * [Using A Function To Solve A Problem](#check-the-maximum-value-of-money-spent-column-d-by-france-column--row-A8-using-the-MAXIFS-function) 
* [Lookup Functions](#Data-Lookup-Functions)
  * [Common Lookup Functions](#Common-Lookup-Functions) 

## Arithmetic Functions
Excel is constantly calculating, formulating, and operating by the use of functions. All functions Excel has to offer use arithmetic to operate, whether it be counting the letters in a word or locating a desired cell. Arithmetic functions are the foundations of several logical functions mentioned below. Harnessing the power of arithmetic functions, Excel elevates the efficiency of data manipulation and analysis, making complex tasks more manageable.

## Common Arithmetic Functions 

**=SUM():** Adds every number in a range of cells, whether it be select cells, rows, or columns. 
```
=SUM(number1,[number2],...) -> =SUM(D:D,I:I) or =SUM(D2,I3)
```
**=AVERAGE():** Calculates the average in a range of cells, whether it be select cells, rows, or columns. 
```
=AVERAGE(number1,[number2],...) -> =AVERAGE(D:D,I:I)
```
**=PRODUCT():** Multiplies a range of cells, whether it be select cells, rows, or columns. 
```
=PRODUCT(number1,[number2],...) -> =PRODUCT(D3, B:B)
```
**=MIN():** Returns the **smallest** number from a range of cells, whether it be select cells, rows, or columns. 
```
=MIN(number1,[number2],...) -> =MIN(D1,A:A)
```
**=MAX():** Returns the **largest** number from a range of cells, whether it be select cells, rows, or columns. 
```
=MAX(number1,[number2],...) -> =MAX(E1,C5:C9)
```
**=SUBTOTAL():** Calculates the subtotal of a range of cells, whether it be select cells, rows, or columns. The subtotal function can choose whether to calculate the min,max,sum,average, product, and many more by the use of selective numbers, (1-9, 10,11,101-111). Using the subtotal function you can calculate a range and **exclude** other subtotals in the range as well as ignoring hidden rows depending on your programmed "function_num". 
```
=SUBTOTAL(function_num,ref1,[ref2],...) -> =SUBTOTAL(9,D1,D3)
```
**AGGREGATE():** Calculates the aggregate calculation of choice from the same choice range as "=SUBTOTAL()" for a range of cells, rows, or columns. Additionally, the "=AGGREGATE()" function allows the ability for "options". This condition allows the "AGGREGATE()" function to ignore certain rows or values depending on the choice range (1-7) that is picked. 
```
=AGGREGATE(function_num,options,ref1,[ref2],...) -> =AGGREGATE(1, 6, A2:A10)
```
## Find the subtotal (SUM) amount of "Money Spent" Column D using the "SUBTOTAL()" function
```
=SUBTOTAL(function_num,ref1,[ref2],...) -> =SUBTOTAL(9,D:D)
```
**Function_num:** This is the number that specifies what function would be used for the subtotal. In this function's case it would be 9, since in the list of functions 9 is the number associated with "SUM"

**Ref1:** This the range of cells that will be used to calculate the subtotal, anywhere from 1-254 references can be used. In the case of the prompt above, the "ref1" would be calculating all of Column D (Money Spent). 

## Logical Operation Functions
Excel hosts several logical functions such as "IF()", "AND()", "OR()", and "NOT()" which are essential tools for data analysis. These functions operate on the basis of true or false conditions, enabling the creation of more advanced, selective functions through the use of merging. Whether used independently or in combination, these functions facilitate a comprehensive examination of data. 

## Common Operation Functions:

**IF():** Evaluates if 1 condition is true / false. Able to customize "TRUE" & "FALSE" statments to be any numerical or text value. 
```
=IF(logical_test,[value_if_true],[value_if_false]) -> =IF(A1 < A2, "TRUE", "FALSE")
```
**AND():** Evaluates if 1 or more conditions are true / false. Unable to customize "TRUE" & "FALSE" statments. 
```
=AND(logical1,[logical2],...) -> =AND(A1 < A2, A2 > A3)
```
**OR():** Evaluates if 1 or more conditions are true / false. If any value statement is TRUE it will show TRUE, if all value statments are FALSE it will show FALSE. Unable to customize "TRUE" & "FALSE" statments.  
```
=OR(logical1,[logical2],...) -> =OR(A2 <> A3)
```
**XOR():** Evaluates if 1 or more conditions are true / false. Returns "TRUE" if only an odd number of conditions are true, if an even number of conditions are true or both are false, it returns "FALSE". Unable to customize "TRUE" & "FALSE" statments.  
```
=XOR(logical1,[logical2],...) -> =XOR(A2 > A3, A3 = A4)
```
**NOT():** Evaluates if 1 condition is true / false, however results are reversed. Unable to customize "TRUE" & "FALSE" statments.
```
=NOT(logical) -> =NOT(A3 > A2)
```
## **Example of a merged logical operation function**
```
=IF(AND(A2 > A3, A4 > A2), "TRUE", "FALSE")
```
## **Common "IF()" Variations**
There are various functions built around "IF()" that enhance their capabilities. Here are several important ones:

**SUMIF():** Adds up values that meet a specific condition. 
```
=SUMIF(Range,Criteria, Sum_range) -> =SUMIF(B:B,B2, D:D)
```
**AVERAGEIF():** Calculates the average of values that meet a specific condition.
```
=AVERAGEIF(Range,Criteria, Average_range) -> =AVERAGEIF(B:B,B2,D:D)
```
**COUNTIF():** Counts the number of cells which hold either text or time that meet a specifiec condition.
```
=COUNTIF(Range, Criteria) -> =COUNTIF(B:B,B2)
```
**IFERROR():** Specifies a value or action to take place incase a formula / function spawns an error message.
```
=IFERROR(value, value_if_error) -> =IFERROR(A1/0, "Cannot divide by 0" ) & =IFERROR(IF(A1/0,"Correct","Incorrect"),"Unable to divide by 0")
```
In summary, "IF()" and its related functions are like a Swiss Army knife for data analysis, allowing you to precisely analyze and manipulate data based on conditions.

## Check the average value of "Money Spent" Column D by "France" Column & Row A8 using the "AVERAGEIF()" function
```
=AVERAGEIF(Range,Criteria, Average_range) -> =AVERAGEIF(A:A,A8,D:D)
```
**Range:** This is the set of cells you want to evaluate. In our example, it's Column A, representing the "Country" column. When selected, it's written as A:A.

**Criteria:** This is the condition that helps narrow down the selection in the range. In our case, the criteria is "France" because we want to find the money spent by "France" in the "Country" column.

**Average_range:** This condition specifies the cells that need to be averaged. In the example, it's Column D, representing the "Money Spent" column.

## Logical Operation Functions "IFS()"
Like the "IF()" function above, "IFS" allow the user to check 1 or more conditions of a formula with more detail. The number of "IFS()" conditions are limitless, as long as the criteria and range can coexist. In the example below the function "MAXIFS()" will be used. 

## **Common "IFS()" Variations**

Like the "IF()" function previously mentioned, there are various functions built around "IFS()" that can elevate their effectiveness. 

All "IF()" functions can become "IFS()", such as "MAXIF()" -> "MAXIFS", "SUMIFS()" -> "SUMIFS()", and "COUNTIF()" -> "COUNTIFS()". While the majority of "IF()" functions can become an "IFS()" function, there are two exceptions:

**MINIFS():** Returns the minimum value in a range of cells based on 1 or more criterias. 
```
=MINIFS(min_range,criteria_range1,criteria1,...) -> =MINIFS(D:D,C:C,C1)
```
**MAXIFS():** Returns the maximum value in a range of cells based on 1 or more criterias. 
```
=MAXIFS(max_range,criteria_range1,criteria1,...) -> =MAXIFS(D:D,C:C,C1)
```
## Check the maximum value of "Money Spent" Column D by "France" Column & Row A8 using the "MAXIFS()" function
```
=MAXIFS(Max_range,Criteria_range1,Criteria1,...) -> =MAXIFS(D:D,A:A,A8)
```
**Max_range:** This is the set of cells that will be analyzed to determine the maximum value regarding the criteria(s) in the form of numerical data. In this case, Max_Range would be Column D, "Money Spent".

**Criteria_range1+:** This is the condition that will help narrow down the selection to a specific column or columns. In this case the criteria_range would be Column A, "Country". 

**Criteria1+:** This condition can be in the form of any criteria, whether it be a numerical, expression, or text value. The Criteria specifies which cells in the Criteria_range will be used to determine the maximum value. In this case the Criteria will be Column A8, "France".

## Data Lookup Functions
Excel features several data lookup functions that allow you to search for specific values in a range or table and retrieve related information. Here are some commonly used data lookup functions in Excel:

## Common Lookup Functions

**SWITCH():** 

**CHOOSE():**

**INDEX():**

**MATCH():**

**=VLOOKUP():**

**=HLOOKUP():**

**XLOOKUP():**

## Text Manipulation Functions

## Date and Time Functions 

