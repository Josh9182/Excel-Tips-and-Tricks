# Excel-Tips-and-Tricks
This a compilation of the most commonly used functions and tricks used frequently to make data analysis just a little bit easier!

## Content
* [Logical Functions](#Logical-Operation-Functions)
  * [Common Variations](#Common-IF-Variations)
  * [Function Example](#check-the-average-value-of-money-spent-column-d-by-france-column--row-a8-using-the-averageif-function)
* [Logical Functions "IFS"](#logical-operation-functions-IFS)
   * [Function Example](#check-the-maximum-value-of-money-spent-column-d-by-france-column--row-A8-using-the-MAXIFS-function)
* [Lookup Functions](#Data-Lookup-Functions)

## Logical Operation Functions
In Excel, logical functions such as "AND()," "OR()," "NOT()," and "IF()" are essential tools for data analysis. These functions operate on the basis of true or false conditions, enabling the creation of more advanced, selective functions through the use of merging. Whether used independently or in combination, these functions facilitate a comprehensive examination of data. 

**Example of a merged logical operation function**
```
=IF(AND(A2 > A3, A4 > A2), "TRUE", "FALSE")
```

## **Common "IF()" Variations**
There are various functions built around "IF()" that enhance its capabilities. Here are three important ones:

**SUMIF():** Adds up values that meet a specific condition.

**AVERAGEIF():** Calculates the average of values that meet a specific condition.

**COUNTIF():** Counts the number of cells that meet a specifiec ondition.
Usual Format:

All these "IF()" functions follow the same structure. You set the range of cells that will be used to generalize the criteria, followed by the more specific cells known as the criteria. Based on whether it's true or false, the function does something specific with the data.

In summary, "IF()" and its related functions are like a Swiss Army knife for data analysis, allowing you to precisely analyze and manipulate data based on conditions.

### Check the average value of "Money Spent" Column D by "France" Column & Row A8 using the "AVERAGEIF()" function
```
=AVERAGEIF(Range,Criteria, Average_range) -> =AVERAGEIF(A:A,A8,D:D)
```
**Range**: This is the set of cells you want to evaluate. In our example, it's Column A, representing the "Country" column. When selected, it's written as A:A.

**Criteria**: This is the condition that helps narrow down the selection in the range. In our case, the criteria is "France" because we want to find the money spent by "France" in the "Country" column.

**Average_range**: This condition specifies the cells that need to be averaged. In the example, it's Column D, representing the "Money Spent" column.

## Logical Operation Functions "IFS()"
Like the "IF()" function above, "IFS" allow the user to check 1 or more conditions of a formula with more detail. The number of "IFS()" conditions are limitless, as long as the criteria and range can coexist. In the example below the function "MAXIFS()" will be used. 

## **Common "IFS()" Variations**

Like the "IF()" function previously mentioned, there are various functions built around "IFS()" that can elevate their effectiveness. 

All "IF()" functions can become "IFS()", such as "MAXIF()" -> "MAXIFS", "SUMIFS()" -> "SUMIFS()", and "COUNTIF()" -> "COUNTIFS()". 

### Check the maximum value of "Money Spent" Column D by "France" Column & Row A8 using the "MAXIFS()" function
```
=MAXIFS(Max_range,Criteria_range1,Criteria1,...)
```

**Max_range**: This is the set of cells that will be analyzed to determine the maximum value regarding the criteria(s) in the form of numerical data. In this case, Max_Range would be Column D, "Money Spent".

**Criteria_range1+**: This is the condition that will help narrow down the selection to a specific column or columns. In this case the criteria_range would be Column A, "Country". 

**Criteria1+**: This condition can be in the form of any criteria, whether it be a numerical, expression, or text value. The Criteria specifies which cells in the Criteria_range will be used to determine the maximum value. In this case the Criteria will be Column A8, "France".

## Data Lookup Functions

