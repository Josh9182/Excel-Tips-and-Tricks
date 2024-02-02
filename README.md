# Excel-Tips-and-Tricks
This a compilation of the most commonly used functions and tricks used frequently to make data analysis just a little bit easier!
## Logical Operation Functions
Functions that include the "IF()" function allow the user to check the conditions of a formula with more detail. 
### Check the Average value of (Money Spent) Column D by (France) Column:Row A8 using the "AVERAGEIF()" function.
```
=AVERAGEIF(Range,Criteria, Average_range) -> =AVERAGEIF(A:A,A8,D:D)
```
The function above lists several parts needed to perform the function correctly. Range, the condition that asks the user to locate the range of cells that will be evaluated. In this example the range would be column A, which is the "Country" column, in which "France" is located in. When selected column A is shown as A:A. Criteria, the condition which narrows down which cell would be selected in the range, whether it be expression, number, or text. In this example, the criteria would be "France" as we are trying to find the money spent by "France", which is located in the "Country" colummn. Lastly, the Average_range, a condition which is asking what are the actual cells which are to be averaged. In this example, the Average_range would be Column D, or "Money Spent". 


IFS(), SWITCH(),
SUMIF(), AVERAGEIF(), COUNTIF(), SUMIFS(), AVERAGEIFS(), COUNTIFS(), MAXIFS(),
MINIFS()) 

