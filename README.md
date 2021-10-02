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


And as expected, since our new refactored code was designed to eliminate the time wasted going down all 313 rows twelve different times, the refactored code executed the analysis more quickly than the original.

![image](https://user-images.githubusercontent.com/90977689/135695482-d724d7c2-2959-449d-a396-95a7e529b68c.png)
![image](https://user-images.githubusercontent.com/90977689/135695629-c65d6420-9f55-4b2f-a1de-07ffd0d58080.png)

Success!

## Summary
This was an interesting exercise that shows how dramatically you can improve runtime with some logical adjustments to the code.  This was a short program where the original code seemed to run pretty quickly so I was surprised at the magnitude of the difference once the code was refactored. At first I struggled quite a bit with this assignment.  Even though it was easy to see the logic change that needed to take place in how the code operated, going from the concept to practical implementation was confusing at first.  It wasn't until I really thought about the relationship between the tickerIndex variable and the output arrays that I was able to come up with a plan.  Being my first time coding, syntax was not intuitive at first but having struggled with it, I feel like I understand it better now than if I were just trying to learn from examples in a book.


### What are the advantages or disadvantages of refactoring code?
I can anticipate several advantages to refactoring code.  The most obvious is the improvement in run time and the elimination of inefficiencies that could be carried through and even compounded in longer code.  The second is the exquisite level of understanding you need to gain of what each line of code is doing in the original code before you can attempt to successfully refactor it.  I thought I understood everything in the original code as I went through the module, but being faced with the need to refactor, it was clear that I needed an even more detailed understanding.  For example, I realize there is additional opportunity to make this code more robust where we increase the ticker volumes.  The challenge instructions and hint for step 3a did not ask for an If Then statement and as a result requires the first row of the analyzed worksheet to already be populated with tickers(0) which is fine if your list is sorted in correct order.  However, for a stock list that is in random order, I think it would be better to build in a check to ensure the ticker symbol matches with tickerIndex using an If Then statement.  Lastly, refactoring provides the opportunity to build additional flexibility into the code.  Since you are redefining the logic, it is possible to make the code easier to adapt to new inputs or output needs in the future.
On the flipside of this, I can imagine that it might also possible to make logic choices that might limit the extensibility of your codeblock to other applications.  In other words, it may get optimized so specifically for the current need that it becomes less useful for future reuse in other scenarios.  Lastly, it does take significant time to refactor code, so there may be a point where the efficiency gains do not merit the amount of time you would need to invest into refactoring in order to achieve those gains.

### How do these pros and cons apply to refactoring the original VBA script?
For this particular challenge, we definitely achieved a significant overall improvement in the runtime.  I also now understand both the original and new code in a level of detail I wouldn't have had otherwise.  Also the refactored code would be the code of choice if I wanted to run this on a much longer list of data.  So those are all pros.  The only con would be that since this specific example is such a short piece of code that the difference between a code running in 1.1 seconds vs 0.17 seconds is probably not meaningful for the enduser.  And of course it took me a long time to figure this challenge out, but it was worth it for the learning.
