# Convert-Column-Name-Values-of-a-.XLSX-file-to-their-Acronym-Values
Project By- Khushvardhan and Jayvardhan Bhardwaj
This Project contains the code as well as sample file which can change the Column Values to their Acronyms.  Along with this it can will write single word column names as it is and will present all the subsection (represented by enclosing them in bracket in original file) with columnname_subsection format. Read the Read.me file to understand the working of whole code.

Libraries Used:-
1.	Openpyxl – Importing this library will help me to access .xlsx file data. 
2.	Pandas – Importing this library will make it easier to perform actions on our data file.
3.	Re- The Python module re provides full support for Perl-like regular expressions in Python.
4.	String – To work on strings using complex libraries.
5.	Itertools – To get combinations of string required to make column names.


Algorithm:-
	SHORT-
While running this program we get the names of each column as we get forward, using itertools. And we save each column name as strings in a list to operate with it.
		LONG/Line by Line –
			- The code iterates through all the strings in a list of size 1.
- It then joins them together to make one string, which is "".- The code is a function that iterates through all the strings in a list.

- The code iterates through all the strings in a list and outputs them as one string, separated by commas.- The code iterates over all the strings in a list and splits them into two lists: one containing only lowercase letters, and another containing uppercase letters.
- The code then iterates through each of these lists, using the row number as an index to find out which column contains that string.
- It then creates a new list with just those values from that column.

- The code starts by creating a list of all the strings in the sheet (the "itertools" library is used for this).
- Then it iterates over each string in turn, splitting it into its lowercase and uppercase parts (using re), finding out which row it's on, and adding that value to cur_row.
- This is done for every string so we can see what rows have been processed already.
- Next, cur_row is split up into two sublists: one containing only lowercase letters (col_count) and one containing uppercase letters (col_count+1).
- For each of these sublists, there are three lines of code: firstly there's an if statement checking whether or not s2 has any capital letters; secondly there's a loop where s2 is taken from range(1, col_count+1- The code attempts to iterate through all the rows of a spreadsheet and replace each row with the next letter in the alphabet.

- The code uses itertools.islice() to create a list of all strings that start with A, then iterates through this list and replaces each string with B, C, D, etc.
			

				
