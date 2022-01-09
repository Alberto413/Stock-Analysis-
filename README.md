# Stock-Analysis-
## Overview: VBA Stock Analysis Project
Using a stock market dataset spanning two years, VBA code was written to calculate total daily volume and rate of return for a year. The stock market data was contained in an Excel spreadsheet and a VBA script was written. The VBA code was refactored in order to make it more efficient to use against larger datasets.

## Purpose:

Refactor the Stock Market Dataset with VBA solution code to loop through all the data one time in order to collect an entire dataset. Then, we’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, we just want to make the code more efficient—by taking fewer steps, using less memory, and improving the logic of the code.

## Background:

> This challenge asks you to edit, and refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.
Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

## Results:

1a). The tickerIndex is set equal to zero before looping over the rows.
> Created a tickerIndex variable and set it equal to zero before iterating over all the rows. Will use this tickerIndex to access the correct index across the four different arrays on VBA Code: the tickers array and the three output arrays created on next requierement.

1b). Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
> Created three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices. In our VBA code, the tickerVolumes array should be a long data type. But in our VBA code the tickerStartingPrices and tickerEndingPrices arrays should be a Variant data type.

2a). Created a for loop to initialize the tickerVolumes to zero. 2b) Loop over all the rows in the spreadsheet

<img width="407" alt="Screen Shot 2022-01-09 at 1 35 46 PM" src="https://user-images.githubusercontent.com/95304774/148695775-6d4d704d-4f38-44c0-8e24-c87115dcb65a.png">

###### The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

> 3a,b,c). if the next row’s ticker doesn’t match, increase the tickerIndex. Created a loop that will loop over all the rows in the spreadsheet. Inside the loop, we created a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker. 
Created an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current closing price to the tickerStartingPrices and tickerEndingPrices variable. 

<img width="514" alt="Screen Shot 2022-01-09 at 1 39 44 PM" src="https://user-images.githubusercontent.com/95304774/148695919-6372b480-adbe-498b-8821-bc986ef4e748.png">

Code for formatting the cells in the spreadsheet is working.
>We make positive returns green and negative returns red, to be a lot easier to determine which stocks did well and which ones didn't. Added some formatting based on the values of the returns.

<img width="514" alt="Screen Shot 2022-01-09 at 2 06 38 PM" src="https://user-images.githubusercontent.com/95304774/148696890-e0e23967-831e-4541-b4a0-aefddb926ad9.png">

Final VBA Analysis and The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png

<img width="278" alt="Screen Shot 2022-01-09 at 2 11 05 PM" src="https://user-images.githubusercontent.com/95304774/148697025-c9e91ef3-9aa1-4f74-8a38-c7092d908fd1.png">

![VBA_Challenge2017](https://user-images.githubusercontent.com/95304774/148694977-63a7f10d-61bc-44ba-a42a-193f963f7eb8.png)


<img width="278" alt="Screen Shot 2022-01-09 at 2 13 03 PM" src="https://user-images.githubusercontent.com/95304774/148697098-4b1c9eda-2c00-419a-aac7-27ab77fdeea0.png">

![VBA_Challenge2018](https://user-images.githubusercontent.com/95304774/148694989-c76a6975-f264-4091-a2af-78aac88835a1.png)

## SUMMARY:

1. What are the advantages or disadvantages of refactoring code?

> You need to perform code refactoring in small steps. Make tiny changes in your program, each of the small changes makes your code slightly better and leaves the application in a working state.

Disadvantages:

> A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions. A complex unstructured code is usually best to split in several functions. Refactoring process can affect the testing outcomes.

Advantages:

> Logical errors easily appear in well structure code that contains nested conditionals and loops.
In our case, using Excel flow displays program logic in a more comprehensible manner, not tied to the order that the underlying code is written.
VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.

2. How do these pros and cons apply to refactoring the original VBA script?

> Improving or updating the code without changing the software’s functionality or external behavior of the application is known as code refactoring. Now, let's think about something, What happens after a couple of days or months yo need to troubleshoot your code? Is it complicated? Is it hard to understand? If yes then definitely you didn’t pay attention to improve your code or to restructure your code.






