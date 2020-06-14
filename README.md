# MS Excel/VBA Project - VBA Wall Street

## Background

![wall-street](images/wall-street.jpg)

For this project, we will use VBA scripting to analyze real stock market data for the 2014-2016 period. 


## Objective

Track performace of stocks for public companies and display results in a simple interface for begginer Excel users.

## Basic Methodology

Script will loop through all the stocks for one year for each run and take the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

 * Highlight positive change in green and negative change in red.
 
 The original dataset is over 25 MB which  

## Trip Duration

For trip duration, there is not enough evidence to state that the trip duration varies significantly with gender. The non-adjusted data includes about between 600 to 700 entries which despite the large magnitude, only comprise less than 1% of the data. Again, users who did not report gender present the largest variation on trip duration, but still no evidence of significant difference.

Finally, start and ending trip locations present some interesting patterns. Male and female riders start travelling from the similar areas, but the destination for males spreads farther apart by comparison. It is unclear why this phenomenon occurs with the information available.  

![trip_duration](images/trip_duration.jpg)

## Preliminary Conclusions

There is not enough evidence to conclude than men are willing to participate in the bike program more so than women based on the data provided. Surveys could serve as a useful tool to complement this study. Since there is no significant difference in trip duration, it would be worthwhile to gather information related to the participant's routine/preferences. For example, female riders may prefer certain areas of the city depending on crime rate. A follow-up assignment would consist on preparing random surveys to gather information about the participant's experience. The data collection system is robust enough and resources should definitely be assign to maintain the program running for the city. 

### Copyright

Arturo Wilson (C) 2020. All Rights Reserved.







### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - Use this while developing your scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.



## Instructions

* Create a script that will loop through all the stocks for one year for each run and take the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* You should also have conditional formatting that will highlight positive change in green and negative change in red.

* The result should look as follows.

![moderate_solution](Images/moderate_solution.png)

### CHALLENGES

1. Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume". The solution will look as follows:
