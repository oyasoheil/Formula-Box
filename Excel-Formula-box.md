# Excel Formula Box
A personal collection of useful Excel formulas I encounter in daily work.
## 2026-05-09 — Extract Text Before a Character
###Use case:### 
Get the text before a specific character (for example before `-` in a product code).
**Formula:**
```excel
=LEFT(A1, FIND("-", A1)-1)

###Use case:###
Clean text that contains extra spaces.
**Formula:**
```excel
=TRIM(A1)
** Explanation **
Removes leading spaces, trailing spaces, and repeated spaces between words.