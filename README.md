# Excel-Tips-and-Tricks
This a compilation of the most commonly used functions and tricks I use frequently to make data analysis just a little bit easier!

## Function Content
### Arithmetic Functions
  * [Common Arithmetic Functions](#Common-Arithmetic-Functions)
      * [Using A Function To Solve A Problem](#Find-the-subtotal-SUM-amount-of-Money-Spent-Column-D-using-the-SUBTOTAL-function)
### Statistical Functions
  * [Merged Function Example](#Example-of-merged-statistical-functions)
      * [Using A Function To Solve A Problem](#Find-the-largest-number-in-Money-Spent-Column-D-without-using-a-MIN-MAX-or-LARGE-function)
### Logical Operation Functions
  * [Common Operation Functions](#Common-Operation-Functions)
      * [Merged Function Example](#Example-of-a-merged-logical-operation-function)
      * [Common "IF()" Variations](#Common-IFS-Variations)
      * [Using A Function To Solve A Problem](#check-the-average-value-of-money-spent-column-d-by-france-column--row-a8-using-the-averageif-function)     
### Logical Operation Functions "IFS()"
  * [Common Operation Functions "IFS()"](#Common-IFS-Variations) 
      * [Using A Function To Solve A Problem](#check-the-maximum-value-of-money-spent-column-d-by-france-column--row-A8-using-the-MAXIFS-function) 
### Lookup Functions
  * [Common Lookup Functions](#Common-Lookup-Functions)
      * [Merged Function Example](#Merged-Function-Examples)
      * [Using A Function To Solve A Problem](#Find-what-city-in-City-Column-B-is-next-to-Canada-in-the-Country-Column-without-using-VLOOKUP-HLOOKUP-or-XLOOKUP-function)
### Text Manipulation Functions
  * [Common Text Functions](#Common-Text-Functions)
      * [Using A Function To Solve A Problem](#The-stakeholder-of-your-project-is-requesting-the-First-Name-Column-A-and-Last-Name-Column-B-to-be-combined-into-one-column-with-a-comma-separating-the-names-Show-the-stakeholder-how-to-combine-these-two-columns-seamlessly-by-the-use-of-a-function)


## Arithmetic Functions
Excel is constantly calculating, formulating, and operating by the use of functions. All functions Excel has to offer use arithmetic to operate, whether it be counting the letters in a word or locating a desired cell. Arithmetic functions are the foundations of several logical functions mentioned below. Harnessing the power of arithmetic functions, Excel elevates the efficiency of data manipulation and analysis, making complex tasks more manageable.

## Common Arithmetic Functions 

**AGGREGATE():** Calculates the aggregate calculation of choice from the same choice range as "=SUBTOTAL()" for a range of cells, rows, or columns. Additionally, the "=AGGREGATE()" function allows the ability for "options". This condition allows the "AGGREGATE()" function to ignore certain rows or values depending on the choice range (1-7) that is picked. 
```
=AGGREGATE(function_num,options,ref1,[ref2],...) -> =AGGREGATE(1, 6, A2:A10)
```
**=AVERAGE():** Calculates the average in a range of cells, whether it be select cells, rows, or columns.
```
=AVERAGE(number1,[number2],...) -> =AVERAGE(D:D,I:I)
```
**MAX():** Returns the largest number from a range of cells, whether it be select cells, rows, or columns.
```
=MAX(number1,[number2],...) -> =MAX(E1,C5:C9)
```
**=MIN():** Returns the smallest number from a range of cells, whether it be select cells, rows, or columns.
```
=MIN(number1,[number2],...) -> =MIN(D1,A:A)
```
**=PRODUCT():** Multiplies a range of cells, whether it be select cells, rows, or columns.
```
=PRODUCT(number1,[number2],...) -> =PRODUCT(D3, B:B)
```
**=SUBTOTAL():** Calculates the subtotal of a range of cells, whether it be select cells, rows, or columns. The SUBTOTAL function can choose whether to calculate the min, max, sum, average, product, and many more by the use of selective numbers (1-9, 10, 11, 101-111).
```
=SUBTOTAL(function_num, ref1, [ref2], ...) -> =SUBTOTAL(9, A:A)
```
**=SUM():** Adds every number in a range of cells, whether it be select cells, rows, or columns.
```
=SUM(number1,[number2],...) -> =SUM(D:D,I:I) or =SUM(D2,I3)
```

## Find the subtotal (SUM) amount of "Money Spent" Column D using the "SUBTOTAL()" function
```
=SUBTOTAL(function_num,ref1,[ref2],...) -> =SUBTOTAL(9,D:D)
```
**Function_num:** This is the number that specifies what function would be used for the subtotal. In this function's case it would be 9, since in the list of functions 9 is the number associated with "SUM"

**Ref1:** This the range of cells that will be used to calculate the subtotal, anywhere from 1-254 references can be used. In the case of the prompt above, the "ref1" would be calculating all of Column D (Money Spent). 

## Statistical Functions
Excel presents several statistical functions which provide tools for data analysis, allowing the user to summarize, calculate, and interpret data. Here are various examples that streamline the data analysis process:

**AVERAGE():** Returns the average value in a range of cells as long as the cells contain numbers. 
```
=AVERAGE(number1,[number2],...) -> =AVERAGE(B2:B10)
```
**AVERAGEA():** Returns the average value in a range of cells. Numbers are evaluated as their numerical value, however the function also evaluates text as set values. Text and "FALSE" data = 0. "TRUE" = 1. 
```
=AVERAGEA(value1,[value2],...) -> =AVERAGE(A2:A10)
```
**COUNT():** Counts the number of cells in a range that contain numbers. 
```
=COUNT(value1,[value2],...) -> =COUNT(E2:20)
```
**COUNTA():** Counts the number of cells in a range that contain any values, whether it be text or numerical data. 
```
=COUNTA(value1,[value2],...) -> =COUNTA(A2:A20)
```
**COUNTBLANK():** Counts the number of empty cells in a range. 
```
=COUNTBLANK(range) -> =COUNTBLANK(A1:E25)
```
**LARGE():** Returns the (k)-th smallest value in the range. (k) is used to show the position from the largest value, if for example the (k) value is "1" then the data that would be shown would be the 1st largest value of the data set.  
```
=LARGE(array,k) -> =LARGE(E2:E25,1)
```
**SMALL():** Returns the (k)-th smallest value in the range. (k) is used to show the position from the smallest value, if for example the (k) value is "1" then the data that would be shown would be the 1st smallest value of the data set.   
```
=SMALL(array,k) -> =SMALL(E2:E25,1)
```
**LAMBDA():** Allows for the creation of a custom function. 
```
=LAMBDA(parameter_or_calculation,[parameter_or_calculation]) -> =LAMBDA(boxes_ordered,boxes_ordered*5)(A14)
```
**LET():** Assigns a name value to a range, text, or number value and allows for the creation of formatted functions. 
```
=LET(name1,value_expression1,[name2],[value_expression2],formula_expression) -> =LET(data,A2:C10,name,A2:A10,occupation,B2:B10,salary,C2:C10,FILTER(data,(occupation = B2))))
```

## **Example of merged statistical functions**

**=SMALL(array,k) + =COUNT(value1,[value2])**: Shows the largest value of the range. 
```
=SMALL(E:E,COUNT(E:E))
```
**=LARGE(array,k) + =COUNT(value1,[value2])**: Shows the smallest value of the range. 
```
=LARGE(E:E,COUNT(E:E))
```
**=LAMBDA(parameter_or_calculation,[parameter_or_calculation]) + =IF(logical_test,[value_if_true],[value_if_false]):** Checks if a number is greater than the value "10".
```
=LAMBDA(x, IF(x > 10, "Greater than 10", "Not greater than 10"))(B1)
```

## **Find the largest number in "Money Spent" (Column D) without using a MIN, MAX, or LARGE function**
```
=SMALL(array,k) + =COUNT(value1,[value2]) -> =SMALL(D:D,COUNT(D:D))
```
**Array:** This is the range of numerical data that will be used to determine the (k)-th smallest value. In the case of this example, the data we are trying to use is Column D, "Money Spent". 

**K:** This is the position from the smallest value in the array that will be returned. 1 = smallest value.  10 = 10th smallest value. In this case, the (k) value would be the function "=COUNT()". "=COUNT()" will find the number of values in the select range.    

**Value1:** This is the range of cells we would like to be calculated. In this example the "value1" is "Money Spent" (Column D), totaling 23 values. When plugged together, The "SMALL()" function will now look at Column D, and with the results from "COUNT()" find the 23rd smallest value, which in this case would be the largest. 

While we can do the "LARGE()" function and obtain the same answer, the ability to have options allows excel to be an extremely versatile and fantastic resource for data analysis. 

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
**=IF(logical_test,[value_if_true],[value_if_false]) + =AND(logical1,[logical2],...):** Returns the value "TRUE" or "FALSE" depending on the logical question asked, in this case "A2 > A3, A4 > A2". 
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
Excel features several data lookup functions that allow for the search of specific values in a range or table. Here are some commonly used data lookup functions in Excel:

## Common Lookup Functions

**CHOOSE():** Based upon a value (index_num), returns select values based upon conditions.
```
=CHOOSE(index_num,value1,[value2],...) -> =CHOOSE(A29,"Purple","Blue","Red")
```
**SWITCH():** Based upon a value (expression), returns select values based upon multiple, more specific conditions.
```
=SWITCH(expression,value1,result1,[value2], [result2]) -> =SWITCH(B32,"Red",10%,B33,"Green",15%)
```
**MATCH():** Returns the row position of an item determined by certain criteria. 
```
=MATCH(lookup_value, lookup_array,[match_type]) -> =MATCH("Josh" or A2,A:A, 0)
```
**INDEX():** Returns a value in a given range based off a particular row and column. 
```
=INDEX(array,row_number,[column_num]) -> =INDEX(A:B,2,1)
```
**OFFSET():** Returns a value in a given range based off a given reference as well as row and column number. 
```
=OFFSET(reference,rows,cols,[height],[width]) -> =OFFSET(A2,0,1)
```
**=VLOOKUP():** Uses a value from the leftmost column of a table, and locates the desired value based off multiple conditions. 
```
=VLOOKUP(lookup_value,table_array,col_index_num,[range_lookup]) -> =VLOOKUP(A2,Table6,FALSE)
```
**=HLOOKUP():** Uses a value from the topmost row of a table, and returns the desired value based off multiple conditions. 
```
=HLOOKUP(lookup_value,table_array,row_index_num,[range_lookup]) -> =HLOOKUP(D2,Table6,3,FALSE)
```
**XLOOKUP():** Uses a value, no matter the location, and returns a value/values from a table by using multiple conditions. 
```
=XLOOKUP(search_key,lookup_range,result_range,[missing_value],[match_mode],[search_mode]) -> =XLOOKUP(A3,A:A,B:B,[Value Not Found])
```
**TRANSPOSE():** Converts a vertical range of cells to a horizontal range, or vice versa.
```
=TRANSPOSE(array) -> =TRANSPOSE(Table6)
```
## Merged Function Examples

**=INDEX(array,row_number,[column_num]) + =MATCH(lookup_value, lookup_array,[match_type]):** Uses a value from the leftmost column of a table, and locates the desired value based off multiple conditions.
```
=INDEX(B:B,MATCH(A2,A:A,FALSE)) 
```
**=OFFSET(reference,rows,cols,[height],[width]) + =MATCH(lookup_value, lookup_array,[match_type]):** Uses a value from the topmost row of a table, and returns the desired value based off multiple conditions. 
```
=OFFSET(A2,MATCH(B2,B:B,0),1)
```
**=IFERROR(value, value_if_error) + =VLOOKUP(lookup_value,table_array,col_index_num,[range_lookup]):** Search for data and replace an "#N/A" error message with desired message. 
```
=IFERROR(VLOOKUP("Canada", A2:E25, 1, FALSE), "Value Not Found")
```
## Find what city in "City" Column B is next to "Canada" in the "Country" Column without using "VLOOKUP()", "HLOOKUP()", or "XLOOKUP()" function.
```
=INDEX(array,row_number,[column_num]) + =MATCH(lookup_value, lookup_array,[match_type]) -> =INDEX(B:B,MATCH(A2,A:A,FALSE),1) 
```
**Array:** This is the range of cells that you would like the end value to be found in. In the case of this example, it will be Column B (B:B) as that is where the value Toronto, from "City" Column B will be. 

**Row_number:** This is the row from the "array" that will be used to return a value. In the case of this example, the "MATCH()" function is used to locate the row number for this function. In order to show results from "row_num" the "column_num" must be implemented. In general, "MATCH()" is used to locate a desired value's row number so implementing into "INDEX()" allows for a fantastic combination.  

**Column_num:** This selects a column from the "array" that will be used to return the desired value. In order to show results from "column_num" the "row_num" must be implemented. In this example, since Column B (B:B) is already implemented as the desired column, the column_num will be 1 as it is 1 more column than A2, the value that match is using to find the row number. 

**Lookup_value:** This is the value used to find the desired value in the array. It can be any value whether it be numerical, text, or logical. In this example the "lookup_value" is A2 as it is the value next to the desired value, the city parallel from "Canada", B2. 

**Lookup_array:** This is the range of cells that contain the "lookup_value". In this case this would be Column A (A:A), as it is the "Country" column which hosts "Canada", A2. 

**Match_type:** This is the number indicating how specific of a result you would like to return. 1,0,-1. 

1 = Less Than Exact

0 = Exact Match

-1 = Greater Than Exact

In this case, the "match_type" for this problem is going to be 0 or -1. Both options would be satisfactory. In general 0 is the primary choice as in most scenarios in order to find exactly what you want, you will pick the "exact match" button.  

## Text Manipulation Functions
Excel offers several powerful text functions that allow you to manipulate, format, and extract information from text strings. Here are some commonly used text functions in Excel:

## Common Text Functions

**CHAR():** Rteurns a character based on a specified code number from the PC
```
=CHAR(Number) -> =CHAR(43) = "+"
```
**CONCAT():** Combines a range of selected cells or values from several cells into one cell. 
```
=CONCAT(text1,text2,[text3],...) -> =CONCAT(B15,", ",B16) = Tokyo, Japan
```
**EXACT():** Returns "TRUE" or "FALSE" based on whether two text strings are exactly the same. 
```
=EXACT(text1,text2) -> =EXACT(B24,B32) = TRUE
```
**FILTER():** Filters a range based off provided conditions.
```
=FILTER(range,condition1,[condition2],...) -> =FILTER(A2:25,E2:E25 > 10)
```
**FIND():** Returns the starting position of one text string within another. 
```
=FIND(find_text,within_text,start,[start_num]) -> =FIND("Bi",B25) = Birmingham (1)
```
**FIXED():** Rounds a number to a specificed number of decimals and returns the result as a text value.
```
=FIXED(number,[decimals],[no_commas]) -> =FIXED(333,333,333.00,,TRUE) = 333333333.00
```
**LEFT():** Returns the number of characters from the start of the **left** side of the text string based off the number chosen. 
```
=LEFT(text, [num_chars]) -> =LEFT(B25,5) = Birmi
```
**RIGHT():** Returns the number of characters from the start of the **right** side of the text string based off the number chosen. 
```
=RIGHT(text, [num_chars]) -> =RIGHT(B25,5) = ngham
```
**MID():** Returns the number of characters from the start of the **middle** side of the text string based off the number chosen.
```
=MID(text, start_num,num_chars) -> =MID(B25,5) = mingh
```
**LEN():** Returns the number of characters in a text string
```
=LEN(text) -> =LEN(A25) = 10
```
**LOWER():** Converts all letters in a text string to **lowercase**.
```
=LOWER(text) -> =LOWER(A25) = birmingham
```
**UPPER():** Converts all letters in a text string to **uppercase**.
```
=UPPER(text) -> =UPPER(A25) = BIRMINGHAM
```
**PROPER():** Capatilizes the first letter of the text values selected. 
```
=PROPER(text) -> =PROPER(A25) = Birmingham
```
**REPLACE():** Replaces part of a text string with a different text string.
```
=REPLACE(old_text,start_num,num_chars,new_text): =REPLACE(B25,1,10,"Paris") 
```
**SUBSTITUTE():** Replaces a text string with a different text string.
```
=SUBSTITUTE(text,old_text,new_text,[instance_num]) = SUSBSITUTE("Birmingham",A25,"Ottawa")
```
**REPT():** Repeats a given text a given number of times.
```
=REPT(text,number_times) -> =REPT("Hello ", 5) = "Hi Hi Hi Hi Hi"
```
**TRIM():** Removes all spaces from a text string except for single spaces between words. 
```
=TRIM(text) -> =TRIM(A25)
```

## The stakeholder of your project is requesting the "First Name" Column A and "Last Name" Column B to be combined into one column with a comma seperating the names. Show the stakeholder how to combine these two columns seamlessly by the use of a function.
```
=CONCAT(text1,text2,[text3],...) -> =CONCAT(A1,", ",B1)
```
**text1:** 

**text2:**

**text3:**
