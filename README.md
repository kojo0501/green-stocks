# Green Stocks

## Overview of Project

The purpose of this project was to create data that Steve can analyze to make informed investments with his parents' money. His parents want to invest in green energy. The first stock they are considering, Daqo, dropped 63% in 2018, so Steve is interested in analyzing other options. The workbook generates data for a dozen stocks. However, the code is refactored for efficiency, without using nested loops, so more data can be added and processed.

## Results

### Using a Nested Loop
The original code used a nested loop which required more computing resources to execute.

![nested loop](https://user-images.githubusercontent.com/24308495/135791158-6d2dd374-1119-41a5-95ed-f6a26869a013.PNG)

In this code, the j-loop is nested in the i-loop as a way to collect data for analysis. As a result, these are the times it takes to execute the code.

2017 with nested loop:
![2017 nested](https://user-images.githubusercontent.com/24308495/135791394-a201bf3b-ca18-4655-8af7-a355f8e72d2c.PNG)

2018 with nested loop:
![2018 nested](https://user-images.githubusercontent.com/24308495/135791411-0927fda9-f18c-46a9-87f7-a55e280c9077.PNG)

Now compare these numbers with the results below for a single loop.

### Using a Single Loop
The refactored code runs much quicker and more efficiently. None of the loops it uses are nested.

![single loop](https://user-images.githubusercontent.com/24308495/135791667-ef7f19e0-6c6a-45ff-bde5-5e6afe15c175.PNG)

This single loop leverages arrays with the tickerIndex variable to gather the same data as the nested loop referenced above. However, it runs much faster.

2017 with single loop:
![VBA_Challenge_2017](https://user-images.githubusercontent.com/24308495/135792687-aaa6a27d-59c4-4bfc-af11-7f274970aeff.PNG)

2018 with single loop:
![VBA_Challenge_2018](https://user-images.githubusercontent.com/24308495/135792771-9b571e4d-5a82-4913-8558-99fbea99601b.PNG)

As you can see by comparison, the refactored code runs about six times quicker, allowing Steve to expand the size of the dataset with fewer limitations than if it had used a nested loop.

## Summary

### Advantages and Disadvantages of Refactoring Code

A primary advantage to refactoring the code is that makes it run faster and requires fewer computing resources. This is not a significant advantage when doing uncomplicated analysis of relatively small data sets, however when using larger data sets it can make the difference between being able to run code on a laptap versus requiring a server to execute code (or simply not being able to execute the code at all).

The disadvantages are that it can be time consuming. The refactored code has to be debugged, and critical code may break in the refactoring process. For large amounts of complicated code, a programmer must be intentional about what changes they intend to make, and how that will impact the rest of the code. The longer the code, the more planning will be required before the code is changed, and the more debugging that will need to occur to make the code run.

### Advantages and Disadvantages of the Original and Refactored VBA Script

The advantage to refactored VBA script is that it run quicker. Steve will be able to add additional market data by expanding the tickers array and, in a few spots, increasing the number of loops to match the number of tickers within the array. A disadvantage to the refactored VBA script was that it took time to write. Without programming the ticker information Steve wants to add to the set, the refactored code will not include it. My belief is that it will need to be refactored an additional time to find an elegant solution for sorting new ticker information and assigning it to the tickerIndex array.

The advantages of the original VBA code was that it seemed easier to understand because it used fewer arrays. The simpler variables easier to conceptualize, and it was easier to explain what was going-on. The process the code was running seemed more transparent. The disadvantage is that the code would have had a lower limit on the amount of data it could process, and as data was added it would have taken longer to process it than the refactored code. Additionally, it also had all the disadvantages that still apply to the refactored code.
