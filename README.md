# expense-manager
Converts expense notes from a simple text file to excel sheet calculating total expense


## History
I like to note down my daily expenses in Google Keep just for the sake of record. 1 note per month is maintained having list of all the expenses done in that month along with dates.
<br>It is one thing to take down notes and another to calculate total daily and monthly expense on that basis which is very boring and time consuming process.
<br>Every other month I thought of shifting these records to an excel sheet for proper maintenance but failed to do so due to laziness.
<br>I was so lazy that 5 years of notes piled up without having any single entry in the excel sheet but at the same time was so consistent that I was always noting down my expenses even though I was not auditing/analyzing my expenses.
<br>At this point it became impossible to create manual entries in excel for this much data (laziness!) and one fine morning I decided that I would be writting a program to convert my monthly expense note into an excel table calculating total daily and monthly expense at the same time.

## Expense format
The format of noting down any expense was consistent so it was easy to decipher the pattern programmatically.
### Format-

<br>[Date]
<br>[Item(s)][Delimiter][Price(s)]
<br>[Item(s)][Delimiter][Price(s)]
<br>..
<br>..
<br>
<br><b>Date -</b> Any format is suitable as while writing in excel it will be written of as string and not date to avoid handling different date formats. Format in my notes - MM/DD, though some earlier notes had DD/MM format as well.
<br><b>Item(s) -</b> Can be any string including symbols.
<br><b>Delimiter -</b> Hyphon (-), hardcoded. Can be made configurable with little effort.
<br><b>Price(s) - [Minus][Number][Plus/Minus][Number]</b>
<br>Can be a simple number or multiple numbers led by +/- sign. Negative expense implies money is credited. For first price, only negative sign is applicable and plus sign is not required, for following prices either of the sign is mandatory.

### Note -
* There can be 'n' number of space between any of the mentioned token (item, delimiter or price) though leading and trailing spaces in each separate token would be truncated.

## Steps
* Copy note content and save in a text file and save, and pass absolute file address to fileName variable in App.java (currently hardcoded as I was replacing content in the same input/output file, can be easily made as input field).
* Pass absolute output file name to outputFileName in App.java (if file already exists then it will be overwritten).
* Run the project and output excel will be created and opened as well for ease of copying into the main sheet.

### Note -
* First line in output excel is intentionally left blank to fill in the month (can also be made as an input)
* Number of rows for each date will be equivalent to the maximum number of item rows on any particular date in that month.
