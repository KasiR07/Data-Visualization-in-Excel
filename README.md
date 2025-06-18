# Imp Formulas for Data Analysis in Excel
iferror() : This is used to replace an error value with another value or leave it blank. To leave it blank: iferror(<original formula>, " "). Shift + right arrow

# String Functions
there are three types: left, mid and right
The syntax is =left(<cell number>, <no.of characters to print>) to select the characters from the left side. Eg; AL 12345 ---> AL . Same shift+down arrow or drag it till down to apply changes for all.
Right works the same way as well. But if we want to format from a specific value then: =mid(<cell number>, s<start value>, <no.of characters>)

# Concatenation
To join to columns or rows, we use the synatx: =concat(<cell number1>,<cell number2>). To make it a little more presenetable, we can also be like: =concat(<cell number1>, ", ", <cell number2>)
Another way of doing this is: =<cell number1>&", "&<cell number2>

# Sequence
To write a sequence of numbers directly with a specified endpoint: =sequence(50) to write the numbers vertically until 50. If we want horizontal, then we can write: =sequence(1,50). But if we 
write(2, 50) then 1-50 will be printed in first column and 51-100 will be printed in second. For printing dates, the syntax is: =sequence(rows, columns, date(yyyy,mm,dd), frequency) ie. 
=sequence(50, , date(2025,01,01), 7) means that 50 rows printed, 0 columns printed(same column printed), start date, frequency on when the dates will repeat)

# E date
This is used to write a start month and carry on the sequence using frequency of months: edate(jan-23, 3) will give error because first we have to write the date in one cell and then reference it: edate(c3,3) will give dates from jan-23, apr 23...

# To select and display Largest or Smallest Value
Max() function will give the largest value BUT To select from one table and print seperately into another table: =large(<staring cell number> : <ending cell number>, <position1> : <last position>). The position basically refers 1 for the largest, 2 for the 2nd largest etc for all value in the table to fill up based on the customised size. Similar for small values as well.

# To select sum of all the repeating values in the table (eg: if 7 is repeating five times, the total value is:)
We use the function for SUMIFS for whom the syntax is: sumifs(sum range, value range, cell number of referred cell). In the above example we are finding out all the values where 7 is repeating as that is the input month given. Hence, the syntax is =SUMIFS(f3:f14, e3:e14,I3) and here I3 is the input cell.
Sometimes, there might be other variables eg: When we are selecting the 'amazon' field, the 'amazon uk', 'amazon inc' weren't being analyzed. Hence, we can change it and write it as:
=SUMIFS(f3:f14, e3:e14,I6&"*") which shows that variables after the text is also included.

# Filter function
This is used to filter the values in the table such that we can use them accordingly. For example, if we are given value (eg:400,000) and we want to classify all the values greater than that, then we can write it as =filter(<table size>,<value range>, <condition>) ie. =filter(B5:C14, C5:C14 > C2) where C2 cell contains the given value. As we change the value in C2, the table changes accordingly.


