# Excel-Tips-and-Tricks
This a compilation of the most commonly used functions and tricks I use frequently to make data analysis just a little bit easier!

## Content
* [Logical Operation Functions](#Logical-Operation-Functions)
  * [Common Operation Functions](#Common-Operation-Functions)
      * [Merged Function Example](#Example-of-a-merged-logical-operation-function) 
      * [Common Variations](#Common-IF-Variations)
      * [Solving Function Problem](#check-the-average-value-of-money-spent-column-d-by-france-column--row-a8-using-the-averageif-function)
* [Logical Operation Functions "IFS"](#logical-operation-functions-IFS)
  * [Function Example](#check-the-maximum-value-of-money-spent-column-d-by-france-column--row-A8-using-the-MAXIFS-function)
* [Advanced Logical Functions](#Advanced-Logical-Functions)
* [Lookup Functions](#Data-Lookup-Functions)

## Logical Operation Functions
In Excel, logical functions such as "IF()", "AND()", "OR()", and "NOT()" are essential tools for data analysis. These functions operate on the basis of true or false conditions, enabling the creation of more advanced, selective functions through the use of merging. Whether used independently or in combination, these functions facilitate a comprehensive examination of data. 

## Common Operation Functions

**IF():** Evaluates if 1 condition is true / false. Able to customize "TRUE" & "FALSE" statments to be any numerical or text value. 
```
=IF(logical_test,[value_if_true],[value_if_false]) -> =IF(A1 < A2, "TRUE", "FALSE")
```
**AND():** Evaluates if 1 or more conditions are true / false. Unable to customize "TRUE" & "FALSE" statments. 
```
=AND(logical1,[logical2]...) -> =AND(A1 < A2, A2 > A3)
```
**OR():** Evaluates if 1 or more conditions are true / false. If any value statement is TRUE it will show TRUE, if all value statments are FALSE it will show FALSE. Unable to customize "TRUE" & "FALSE" statments.  
```
=OR(logical1,[logical2]...) -> =OR(A2 <> A3)
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
There are various functions built around "IF()" that enhance its capabilities. Here are several important ones:

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
=MINIFS()
```
**MAXIFS():** Returns the maximum value in a range of cells based on 1 or more criterias. 
```
=MAXIFS()
```
## Check the maximum value of "Money Spent" Column D by "France" Column & Row A8 using the "MAXIFS()" function
```
=MAXIFS(Max_range,Criteria_range1,Criteria1,...)
```

**Max_range:** This is the set of cells that will be analyzed to determine the maximum value regarding the criteria(s) in the form of numerical data. In this case, Max_Range would be Column D, "Money Spent".

**Criteria_range1+:** This is the condition that will help narrow down the selection to a specific column or columns. In this case the criteria_range would be Column A, "Country". 

**Criteria1+:** This condition can be in the form of any criteria, whether it be a numerical, expression, or text value. The Criteria specifies which cells in the Criteria_range will be used to determine the maximum value. In this case the Criteria will be Column A8, "France".

## Advanced Logical Functions

While operation functions utilize data to return true or false values, logical functions instead work with data to return numerical or text values based on the functions evaluations.  

## Common Logical Functions
There are various logical functions used to evaluate data, these include:

**SWITCH()**:

**CHOOSE()**:

## Data Lookup Functions

