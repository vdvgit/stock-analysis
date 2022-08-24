# stock-analysis
## Overview of Project
Edit (refactor) the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, determine whether refactoring your code successfully made the VBA script run faster. 

## Results
- Execution times for the refactored code for 2017 and 2018: 
![Alt text](https://raw.githubusercontent.com/vdvgit/stock-analysis/main/resources/VBA_Challenge_2017.png) ![Alt text](https://raw.githubusercontent.com/vdvgit/stock-analysis/main/resources/VBA_Challenge_2018.png)

- 2017/2018 execution time of the refactored code was slightly better than my original code (0.3 - 0.4 second difference). 

- I was able to do this by looping the entire sheet (vs specific cells like in the original code): 
![Alt text](https://raw.githubusercontent.com/vdvgit/stock-analysis/main/resources/loop_example.png)

## Summary
1. What are the advantages or disadvantages of refactoring code?
  - Advantage: completes quicker, cleaner/easier to read/debug. 
  - Disadvantage: takes longer to write. 

2. How do these pros and cons apply to refactoring the original VBA script?
  - This is a simple example, so the time to re-write the code and subsequent time savings during execution are negligible. 
