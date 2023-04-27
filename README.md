# VBA-Challenge
Module 2 Challenge - VBA

For this assignment, I began by setting a variable for the worksheets, and then started a loop for all the worksheets, as well as activating them. I found that without 'ws.Activate' the loop would not run through all worksheets at once. The beginning of the loop sets variables for the ticker names, total stock per ticker, yearly change per ticker (as well as declaring a variable to hold the initial open price) and percent change per ticker, as well as keeping track of each ticker and finding the last row in the master list.

Now we can begin to loop through all the tickers, so I began a nested loop that would find if we moved to a different ticker name or not, the yearly change, percent change, and then move to the next ticker's open price. This nested loop also adds up the entire stock total for one ticker, prints these values into a new table, and colors positive and negative change for appropriate cells (conditional formatting) on the same worksheet, and this ends the nested loop.

To summarize the findings, I declared variables for greatest % increase, decrease, and total volume, and began another nested loop (still within the initial loop for all worksheets, but outside the first nested loop) to find these values as well as their corresponding ticker names. This ends the second loop.

The last part of code prints headers (columns and rows) for the summary table of statistics and autofits all columns so they are easier to read.

This ends the whole loop for all worksheets.

Notes: Starting out, I initially used dblMin and dblMax to find both the Greatest % Increase and Greatest % Decrease in my second table of results, but I found it would be easier to loop through the first table I created to find both those values and their corresponding tickers. During my initial trial for these values I found this website:
https://www.excelanytime.com/excel/index.php?option=com_content&view=article&id=105:find-smallest-and-largest-value-in-range-with-vba-excel&catid=79&Itemid=475
that helped me find the maximum and minimum values in the column. I ended up not needing these functions though!
I also received some advice from both my classmates Ethan Musa and Sneha Thomas about how to declare and store my opening and closing prices as variables. I think I was trying to run my loop without doing this and it took me a little bit for it to finally click!
