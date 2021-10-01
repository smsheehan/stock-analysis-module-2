# Stock Analysis - Module 2 Challenge

## Overview of Project
Module 2 was an introduction to visual basic for applications (vba) and was focused on manipulation of the 2017 and 2018 stock information from 12 different environmentally companies.  These "green stocks" needed to be analyzed to help our fictional friend "Steve" who was interested in providing his parents with an analysis to help them make investment decisions.  Throughout the module, in the excel developer environment, we built vba code that enabled the end user to select a year and then receive as output the Total Daily Volume and Rate of Return for the year for each stock presented in a conditional formatted table.  Additionaly we learned to incorporate a run timer so we could see how long it took our code to execute.  The purpose of this challenge was for us to refactor the code to make it more efficient in how it processes data and as a result run the process more quicly than the original code.

## Results
Initial review of the code built during the module revealed that the primary source of inefficiency was that it required a check of every line in the spreadsheet to see if the first cell matched the first indexed item in our ticker array.  If it was a match, the code then performed a nested loop.  If it was not a match, it would move to the next row.  Then after it ran through all the rows, it would start at the top and look in every row to see if there was a match to the second indexed item in our ticker array. It would go through this for all of the 12 indices in the ticker array.  In essence the primary loop of the code required checking 313 rows times 12 tickers or 3,756 evaluations.  Additionally, before looping to the next ticker, it would directly output the values for that ticker into the "All Stocks Analysis" worksheet before executing "Next i".  Clearly from this analysis, we can gain efficiency by refactoring our code to go through the rows once and collecting all the relevant information for that row.  Additionally, populating the table all at once at the end should also be a slight efficiency gain.  These are the runtimes for the original code:

![image](https://user-images.githubusercontent.com/90977689/135681548-52064035-2846-4773-982c-ddac26b1f5eb.png)
![image](https://user-images.githubusercontent.com/90977689/135681636-307d6c63-24b6-46f9-8cce-4b6c85ab5be4.png)


The first step in refactoring the code is to create a mechanism for evaluating the ticker column cell for which ticker symbol is contained in the cell so that the rest of the information from the row analysis can be put in the appropriate place in our output arrays (which we also need to create).  First I created a variable named tickerIndex and set the value to zero:
![image](https://user-images.githubusercontent.com/90977689/135694564-95ca8e76-830d-4c5b-950b-9dbeb4bf65e1.png)
Then I created the output arrays:
![image](https://user-images.githubusercontent.com/90977689/135694649-f36895d4-dcd1-4083-bd66-d8c2ea0049d2.png)

After creating a for loop to set all tickerVolumes in the array to zero, I set up the for loop which enables us to extract all of the key information from the row.  The key here is using the tickerIndex as the way to set where in each array a value will be placed:  
![image](https://user-images.githubusercontent.com/90977689/135694852-54e558d2-8744-4c80-a2e8-d464bbea4a19.png)

Then I created code to output each of the arrays into the "All Stocks Analysis" worksheet This was done with a simple for loop:
![image](https://user-images.githubusercontent.com/90977689/135695253-8fbda971-3c9b-41b8-8c31-d248760724b4.png)

After standard formatting, the code was ready to be run.
The refactored code delivered the same spreadsheet results as the original code.
![image](https://user-images.githubusercontent.com/90977689/135662302-b9761ef5-e525-4597-8e0c-789a5788d4ee.png)

  ![image](https://user-images.githubusercontent.com/90977689/135662046-c6da964f-107b-4918-88ac-260a8cd8f708.png)
The refactored code executed the analysis more quickly than the original.

## Summary

### What are the advantages or disadvantages of refactoring code?


### How do these pros and cons apply to refactoring the original VBA script?
