# Excel-Tips-and-Tricks
This a compilation of the most commonly used functions and tricks used frequently to make data analysis just a little bit easier!

## Content
* [Logical Functions "IF"](#Logical-Operation-Functions-IF)
  * [Common Variations](#Common-IF-Variations)
  * [Function Example](#check-the-average-value-of-money-spent-column-d-by-france-column--row-a8-using-the-averageif-function)
* [Logical Functions "IFS"](#logical-operation-functions-IFS)

## Logical Operation Functions "IF()"
The "IF()" function is crucial in data analysis. It helps provide different outcomes based on whether a condition is true or false. You can use it alone or combine it for a more comprehensive examination. 

### **Common "IF()" Variations**
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
Like the "IF()" function above, "IFS" allow the user to check conditions of a formula with more detail. However, besides 1 condition, it can be ∞, as long as the criteria and range can coexist. In the example below it will be shown through "MAXIFS()". 
