# Stock Analysis

## Overview of Project

### Purpose
Steve has been tasked with helping his parents who are interested in investing in green energy.  They've already decided to invest all their money into a single stock (DAQO New Energy Corporation) without doing much research, and Steve is concerned about diversifying their funds.  Steve has collected the data of DAQO and other green energy stocks, and has asked us to help him analyze all of it.  This project is designed to assist him by automating analyses through coding so he can reuse this code on any of the stocks and minimize chances of errors.

## Results

![VBA_Challenge_2017.PNG](https://github.com/mathur-nikita/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

![VBA_Challenge_2018.PNG](https://github.com/mathur-nikita/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

## Summary

1) What are the advantages or disadvantages of refactoring code?

####Advantages 
- One of the advantages to refactoring code is not re-inventing the wheel.  Once the logic to solve a particular type of problem has been created, that same logic can be applied to other calculations that require the same type of input and results.  You don't have to come up with the steps to solve the same problem again.

- Another advantage to refactoring code is efficiency. If the code to solve a particular type of problem or calculation exists and there are multiples times it will need to be solved, that code can be adjusted to handle input that isn't specific to only one of the calculations.  The rest of your code can access this one piece of logic in one location instead of having it duplicated in multiple locations across your code.

- Another advantage is making changes in one location.  If the requirements for the calculations being made or steps to take have to be changed, there's only one instance of code that would need to be adjusted to meet new criteria and that new logic will apply to all calculations.

#### Disadvantages
- A disadvantage to refactoring code is understanding what changes need to be made to the original logic.  Code might have values that are specific/hard-coded throughout and sometimes the logic has to be re-adjusted to accomodate for th removal of hard-coded values.

- Another disadvantage can be the amount of time it takes to refactor code.  It isn't always clear how much code actually needs to be changed and how many times it needs to be changed, and this can pose a challenge when there's a time limit involved.

2) How do these pros and cons apply to refactoring the original VBA script?
