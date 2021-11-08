# Stock Analysis

## Overview of Project

### Purpose
Steve has been tasked with helping his parents who are interested in investing in green energy.  They've already decided to invest all their money into a single stock (DAQO New Energy Corporation) without doing much research, and Steve is concerned about diversifying their funds.  Steve has collected the data of DAQO and other green energy stocks, and has asked us to help him analyze all of it.  This project is designed to assist him by automating analyses through coding so he can reuse this code on any of the stocks and minimize chances of errors.

## Results

### 2017 Overview

![VBA_Challenge_2017.PNG](https://github.com/mathur-nikita/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

In 2017, it can be seen the return for nearly all of the green stocks was positive, with DQ having the highest percentage of return.  Only TERP had a negative return.  From this set of stocks it would seem 2017 was a good year for green stocks overall.  

The time required to make all these calculations took about 0.2 seconds.  In comparison, the old version of the script took about 1.1 seconds and did not include any output formatting.

![VBA_Challenge_old_2017_results.png](https://github.com/mathur-nikita/stock-analysis/blob/main/Resources/VBA_Challenge_old_2017_results.png)

---
### 2018 Overview

![VBA_Challenge_2018.PNG](https://github.com/mathur-nikita/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

In 2018, things were much more different.  It can be seen that most green stocks had a negative return, with DQ having the most negative percentage of return.  Only ENPH and RUN had positive returns.  From this set of stocks it would seem 2018 was a poor year for green stocks overall.

The time required to make all these calculations took about 0.2 seconds.  In comparison, the old version of the script took about 1.3 seconds and did not include any output formatting.

![VBA_Challenge_old_2018_results.png](https://github.com/mathur-nikita/stock-analysis/blob/main/Resources/VBA_Challenge_old_2018_results.png)

---

### Some thoughts:
- DQ seems to be the most volatile green stock to invest in.  You can make very high returns if the year is good for green stocks, but you can lose a lot too if the year is poor for green stocks.
- ENPH and RUN seem to be more promising as green stocks that can result in positive returns even during a poor year, with ENPH being the strongest candidate because of its stability.
- It would seem that there isn't a balance of positive and negative returns for green stocks; either the majority follow a positive return trend or a negative return trend.  While diversification might not save Steve's parents from incurring losses, it can probably help minimize them during a poor year (ex. about half of the green stocks shown for 2018 had losses under 25%).  

## Summary

### What are the advantages or disadvantages of refactoring code?

#### Advantages 
- One of the advantages to refactoring code is not re-inventing the wheel.  Once the logic to solve a particular type of problem has been created, that same logic can be applied to other calculations that require the same type of input and results.  You don't have to come up with the steps to solve the same problem again.

- Another advantage to refactoring code is efficiency. If the code to solve a particular type of problem or calculation exists and there are multiples times it will need to be solved, that code can be adjusted to handle input that isn't specific to only one of the calculations.  The rest of your code can access this one piece of logic in one location instead of having it duplicated in multiple locations across your code.

- Another advantage is making changes in one location.  If the requirements for the calculations being made or steps to take have to be changed, there's only one instance of code that would need to be adjusted to meet new criteria and that new logic will apply to all calculations.

#### Disadvantages
- A disadvantage to refactoring code is understanding what changes need to be made to the original logic.  Code might have values that are specific/hard-coded throughout and sometimes the logic has to be re-adjusted to accomodate for the removal of hard-coded values.  It's also possible that a change might be missed if there are multiple lines that need to be updated, and it will be challenging to track that missed instance later.

- Another disadvantage can be the amount of time it takes to refactor code.  It isn't always clear how much code actually needs to be changed and how much time will need to be spent making each change, and this can pose a challenge when there's a time limit involved.  It would need to be determined if the time and effort required to make these changes is worth it (ex. will there be a noticeable improvement to the results?) or if it's just "busy work" and our time would be better spent elsewhere.

### How do these pros and cons apply to refactoring the original VBA script?

In a very early version of our code (from early on in the module) we were hard-coding the value of the year we wanted to analyze data for, and that hard-coded value was used to activate the corresponding worksheet.  Any time we needed to switch to a different year, we would have to change the code to use a new value for the year.  The refactored version of the script now allows the user to input the year they are interested in analyzing, and on the back end the code is able to activate the correct worksheet based on that input.  There is no need to keep changing the values of variables and no need to have duplicates of the script for all possible years.

In an earlier version of our script a nested for loop was being used to iterate through the list of tickers and also iterate through all of the rows in the worksheet.  In the refactored version there is only one for loop used to iterate not only through the list of tickers but also additional lists that hold output values.
  - The "ticker" variable held the name of the ticker as a string and was populated by the outer loop of a nested for loop.  This was replaced with "tickerIndex" which holds the index of a ticker's location in a list as an integer.
  - The "totalVolume" variable held a calculated sum of volumes belonging to a specific ticker (using the "ticker" variable).  This was replaced with an external array called "tickerVolumes" that holds all of the calculated sums.
  - The worksheet of the specific year used to be activated at the beginning of the outer for loop to capture and calculate data, and then at the end of the outer for loop the analysis worksheet would be activated so output could be stored there.  This was replaced with two separate for loops that split the calculation tasks and output tasks.  The year worksheet is activated once before calculations begin, and then the analysis worksheet is activated for the second for loop to complete all outputs.
  - The older code did not handle formatting of the output worksheet; that was left for a different subroutine.  This version of the code handles outputformatting at the end.

The aforementioned changes have affected the readability and efficiency of the code.  It's somewhat easier to follow what's going on because functionality has been separated between doing calculations/capturing data and displaying output.  By storing calculations and data in separate arrays we now have the opportunity to display that output in multiple ways by accessing the same source instead of needing to loop through data multiple times.  We are also able to iterate through the output arrays to perform additional calculations if necessary.  We can also see how much faster the code runs because we've eliminated the use of nested for loops.

The challenging part about this would be taking the time to rework the logic.  The previous version of the script worked and was easy to follow since there were specific variables to hold specific values that could be reset per loop.  The current version of the script goes beyond just presenting calculations; it also provides flexibility should analysis criteria change in the future.  Not only is it necessary to accomodate for this flexibility it's also important to not alter the existing correct functionality.  In our case we were provided with a template that helped guide us to the new solution, but it would have taken much more time had we needed to complete this from scratch.
