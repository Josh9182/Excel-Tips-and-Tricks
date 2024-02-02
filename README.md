# Excel-Tips-and-Tricks
This a compilation of the most commonly used functions and tricks used frequently to make data analysis just a little bit easier!
## Logical Operation Functions "IF()"
Functions that include the "IF()" function allow the user to check the conditions of a formula with more detail. In the example below it will be shown through "AVERAGEIF()", however every "IF()" function is the exact same format besides the function type. 


**IF() Functions**: SUMIF(), AVERAGEIF(), and COUNTIF(). 


### Check the Average value of (Money Spent) Column D by (France) Column:Row A8 using the "AVERAGEIF()" function.
```
=AVERAGEIF(Range,Criteria, Average_range) -> =AVERAGEIF(A:A,A8,D:D)
```
**Range**: This is the set of cells you want to evaluate. In our example, it's Column A, representing the "Country" column. When selected, it's written as A:A.

**Criteria**: This is the condition that helps narrow down the selection in the range. In our case, the criteria is "France" because we want to find the money spent by "France" in the "Country" column.

**Average_range**: This condition specifies the cells that need to be averaged. In the example, it's Column D, representing the "Money Spent" column.

## Logical Operation Functions "IFS()"
Like the "IF()" function above, "IFS" allow the user to check conditions of a formula with more detail. However, besides 1 condition, it can be ∞, as long as the criteria and range can coexist. In the example below it will be shown through "MAXIFS()". 
