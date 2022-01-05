# Refactoring VBA Code

## Overview of Project

Steve asked us to develop VBA code to analyze green stocks performance for years 2017 and 2018 from an Excel data set.  Although this code perfectly runs for the stocks Steve has picked, he wants to analyze more stocks.

So, after finally generating the code, we refactored it so it can run faster to analyze hundreds of stocks.

## Results

The following graphs compare both runs for the original and the refactored code for both years 2017 and 2018.

| 2017 Run for original code | 2017 for refactored code |
|-------|-------|
| ![2017 run original](/resources/Original_2017.png)  | ![2017 run refactored](/resources/VBA_Challenge_2017.png) |




## Summary

In the case of the stock analysis code, evidently the objective was accomplished as the running time was reduced significatively. Refactoring the code led to reducing lines of code and the structure itself turned into a cleaner version so it can be easier reviewed.

When refactoring code we can find ourselves with advantages and disadvantages such as:

**Advantages**
    - Of course, the main advantage is that the final code will run faster and smoothly.
    - We can optimize loops or amount of coding by the use of general variables.
    - Antoher one is that we can find some bugs or logical errors that are not evident in the first run because the code may run without errors until specific cases.
    - We can make the structure cleaner and add comments so further reviews by peers can be easier.

**Disadvantages**
    - I think that the main disadvantage of refactoring is the time employed in going through all the code. Especially if it is long enough.
    - It can derive into complexity instead of just the original objective (i.e. making it cleaner, faster running, etc.).
    - One could get lost in the changes done after a lot of refactoring. Hoping comments will help with this.



Regarding our code refactoring, the main changes to the code were taking out the *Tickets* loop from the *Rows* loop, and creating array variables for the volume and prices. Also reviewing-wise, these changes create an easier way to review the code.

On the cons side I did found myself investing a lot of time because I had an error in the array dimensioning so I had to debug the code several times to find it.