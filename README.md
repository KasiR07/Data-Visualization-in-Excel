
# 📊 Imp Formulas for Data Analysis in Excel

## ✅ Error Handling
- **`IFERROR()`**: Replaces error values with a specified value (or blank).  
  **Syntax:** `=IFERROR(original_formula, " ")`  
  Use this to avoid #N/A, #DIV/0 errors in calculations.

  explaination: iferror() : This is used to replace an error value with another value or leave it blank. To leave it blank: iferror(<original formula>, " "). Shift + right arrow


---

## 🔤 String Functions
- **`LEFT()`**, **`MID()`**, **`RIGHT()`**: Extract specific parts of text.
  - `=LEFT(A1, 2)` → First 2 characters from the left.
  - `=RIGHT(A1, 3)` → Last 3 characters from the right.
  - `=MID(A1, 4, 5)` → 5 characters starting from 4th character.
 
  - explaination: There are three types: left, mid and right
The syntax is =left(<cell number>, <no.of characters to print>) to select the characters from the left side. Eg; AL 12345 ---> AL . Same shift+down arrow or drag it till down to apply changes for all. Right works the same way as well. But if we want to format from a specific value then: =mid(<cell number>, s<start value>, <no.of characters>)


---

## 🔗 Concatenation
- Join text from multiple cells:
  - `=CONCAT(A1, ", ", B1)`
  - Or using `&`: `=A1 & ", " & B1`
 
  - explaination:To join to columns or rows, we use the synatx: =concat(<cell number1>,<cell number2>). To make it a little more presenetable, we can also be like: =concat(<cell number1>, ", ", <cell number2>). Another way of doing this is: =<cell number1>&", "&<cell number2>

---

## 🔢 Sequence Function
- `=SEQUENCE(50)` → Numbers 1–50 vertically.
- `=SEQUENCE(1, 50)` → Numbers 1–50 horizontally.
- `=SEQUENCE(2, 50)` → First column: 1–50, Second: 51–100
- For dates:  
  `=SEQUENCE(50, , DATE(2025,1,1), 7)` → 50 dates incremented weekly from Jan 1, 2025

    explaination:To write a sequence of numbers directly with a specified endpoint: =sequence(50) to write the numbers vertically until 50. If we want horizontal, then we can write: =sequence(1,50). But if we write(2, 50) then 1-50 will be printed in first column and 51-100 will be printed in second. For printing dates, the syntax is: =sequence(rows, columns, date(yyyy,mm,dd), frequency) ie. =sequence(50, , date(2025,01,01), 7) means that 50 rows printed, 0 columns printed(same column printed), start date, frequency on when the dates will repeat)

  

---

## 📆 `EDATE()` Function
- Returns a date N months before/after a start date.
- `=EDATE(C3, 3)` → 3 months after the date in cell C3

-   explaination: This is used to write a start month and carry on the sequence using frequency of months: edate(jan-23, 3) will give error because first we have to write the date in one cell and then reference it: edate(c3,3) will give dates from jan-23, apr 23...


---

## 🔝 Select Largest/Smallest Values
- `=LARGE(range, k)` → k-th largest
- `=SMALL(range, k)` → k-th smallest
- Example: `=LARGE(A1:A10, {1,2,3})` → Top 3 values

-   explaination: Max() function will give the largest value BUT To select from one table and print seperately into another table: =large(<staring cell number> : <ending cell number>, <position1> : <last position>). The position basically refers 1 for the largest, 2 for the 2nd largest etc for all value in the table to fill up based on the customised size. Similar for small values as well.

---

## ➕ Summing with Conditions
- `SUMIFS(sum_range, criteria_range, criteria)`
- Example:  
  `=SUMIFS(F3:F14, E3:E14, I3)`  
  To include partial text matches:  
  `=SUMIFS(F3:F14, E3:E14, I6 & "*")`

  explaination: We use the function for SUMIFS for whom the syntax is: sumifs(sum range, value range, cell number of referred cell). In the above example we are finding out all the values where 7 is repeating as that is the input month given. Hence, the syntax is =SUMIFS(f3:f14, e3:e14,I3) and here I3 is the input cell. Sometimes, there might be other variables eg: When we are selecting the 'amazon' field, the 'amazon uk', 'amazon inc' weren't being analyzed. Hence, we can change it and write it as:
=SUMIFS(f3:f14, e3:e14,I6&"*") which shows that variables after the text is also included.


---

## 🔍 Filter Function
- `=FILTER(range, condition)`
- Example:  
  `=FILTER(B5:C14, C5:C14 > C2)` → Filters values greater than C2

  explaination: This is used to filter the values in the table such that we can use them accordingly. For example, if we are given value (eg:400,000) and we want to classify all the values greater than that, then we can write it as =filter(<table size>,<value range>, <condition>) ie. =filter(B5:C14, C5:C14 > C2) where C2 cell contains the given value. As we change the value in C2, the table changes accordingly.

---


## 🔎 Lookup Functions

### XLOOKUP (Modern & Flexible)
```excel
=XLOOKUP(H4, Table2[Salesperson], Table2[Commission], "Name not found")
```

### VLOOKUP (Vertical Lookup)
```excel
=IFERROR(VLOOKUP(H4, Table2, 2, FALSE), "Name not found")
```

### HLOOKUP (Horizontal Lookup)
```excel
=IFERROR(HLOOKUP(H4, Table2, 2, FALSE), "Name not found")
```

---

## 🔄 Difference Between Lookup Functions

| Feature               | XLOOKUP                         | VLOOKUP                          | HLOOKUP                         |
|-----------------------|----------------------------------|----------------------------------|----------------------------------|
| Lookup Direction       | Vertical & Horizontal           | Vertical Only                    | Horizontal Only                  |
| Approx./Exact Match    | Both (default: exact)           | Both                             | Both                             |
| Returns from...        | Any direction                   | Only right of lookup column      | Only below lookup row            |
| Error Handling         | Built-in                        | Requires IFERROR                 | Requires IFERROR                 |
| Availability           | Excel 365, Excel 2019+          | All Excel versions               | All Excel versions               |

📌 **When to Use:**
- Use **XLOOKUP** for modern, flexible lookups (recommended).
- Use **VLOOKUP** if data is vertical and you're using older Excel.
- Use **HLOOKUP** for horizontal table structures.

- explaination: As the name suggests, we are basically looking for the value in the table and the syntax is =xlookup(<input cell>, <range to find the input>, <value range of input cell) . So if the input value is given in cell H4 as sai, we are searching for his commisiion, then: =xlookup(H4, F3:F24, E3:E24) which gives out the commission of Sai. As we change the input value, that value also changes.


---

## 🧠 INDEX + MATCH Combo

### When to Use:
- More powerful than VLOOKUP when you need:
  - To search **left of** lookup column
  - **Dynamic row and column referencing**

### Syntax:
```excel
=INDEX(result_range, MATCH(row_lookup, lookup_column, 0), MATCH(col_lookup, lookup_row, 0))
```

- `INDEX()` returns a value from a range based on row/col position
- `MATCH()` finds the row or column number of a value
- `0` ensures an exact match

- explaination: These are two seperate functions that are used to analyze both the rows and coulmns to dynamically find out data. First the index() command is used to select the whole data folloed by match() command for specific constraints. 0 is used to indicate the exact precise value whereas 1 indicates . Hence, the syntax will be index(all values, match(given input, all input values, 0), match(given input 2, all input values, 0))


---


