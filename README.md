# Stock Market Analysis with VBA

## Overview of Project

### Purpose
Steve wants to expand and do a little more research for his parents. Instead of analyzing several stocks, he wants to be able to easily analyze thousands of stocks across the entire stock market over the past few years. The ultimate purpose of this challenge will be to refactor/edit the existing code we used for module 2 to determine whether the VBA script will execute faster or not as a result. 
## Results
### Stock Performance 2017
Overall, all stocks in 2017 performed exceptionally well with positive returns at year end with the sole exception of one stock (TERP), which returned a -7% return for the year. 
### Stock Performance 2018
Conversely, all stocks in 2018 performed quite terribly with only two stocks returning a positive return (ENPH and RUN). 
### Execution Time Original Script
The execution time for running the refactored script on the stock market in 2017 was 2.85 seconds.![Original_script_2017](https://user-images.githubusercontent.com/70483866/93010399-abb2cf80-f551-11ea-8938-a791bc723c8a.PNG)

The execution time for running the refactored script on the stock market in 2018 was 1.83 seconds.![Original_script_2018](https://user-images.githubusercontent.com/70483866/93010413-d2710600-f551-11ea-8668-908a339fe26d.PNG)

### Execution Time Refactored Script
The execution times for the refactored script did in fact run faster than the original script as you might have expected. I was able to run the execution times for both years with the following code:  

yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer
     
     endTime = Timer
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

The execution time for running the refactored script on the stock market in 2017 was 2.41 seconds.![VBA_Challenge_2017](https://user-images.githubusercontent.com/70483866/93010327-fe3fbc00-f550-11ea-9e6b-343c767fed5f.PNG)

The execution time for running the refactored script on the stock market in 2018 was 1.27 seconds.  
![VBA_Challenge_2018](https://user-images.githubusercontent.com/70483866/93010355-4c54bf80-f551-11ea-9789-c011a23a6707.PNG)


## Summary

- What are the advantages or disadvantages of refactoring code?

Some advantages include more logical, clear, and succinct code by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. It also is a good way to find bugs in the code and improve the design of the software. Some disadvantages may be that it can be time consuming and if it isn't documented well with comments for future coders, then it can be difficult to follow and edit properly.

- How do these pros and cons apply to refactoring the original VBA script?

It definitely helped the transition to refactoring the original VBA script by having documented and commented/labeled the various sections of code. When it came to refactoring, I could easily reference the comments and headings for guidance and to keep the script more organized and succinct. Refactoring helped turn the tide a bit with exectuion times of running the code in making the process faster and more efficient. It was a bit time consuming and I found myself re-reading the refactored code repeatedly to ensure there were no mistakes and that each line of code made complete sense to both me and future coders.

