# MS Excel/VBA Project - VBA Wall Street

## Background

![Citi-Bikes](images/citi-bike-station-bikes.jpg)



## Objective

Find relationships or patterns in the data based on gender by using New Yor City, one of the largest metropolitan areas in the world, as reference to verify these findings

## Basic Demographics

The dataset from 2019 contains trip information including start and end location, duration, bike used, distance travelled, and rider gender. The dataset contains 404,947 records. Refer to the pie chart to the left for breakdown on gender distribution. Clearly, males use bike as mean of transportation more than women

However, this proportion alone is not good to explain the difference in participation. The data obtain shows an age breakdown that is remarkably similar for both genders considering the difference in population. It's worth noting that riders who did not report their gender are typically older than  half of the users. Notice the data is skewed meaning that most riders tend to be 45 years old or older regardless of gender.

![demographics](images/demographics.jpg)

## Trip Duration

For trip duration, there is not enough evidence to state that the trip duration varies significantly with gender. The non-adjusted data includes about between 600 to 700 entries which despite the large magnitude, only comprise less than 1% of the data. Again, users who did not report gender present the largest variation on trip duration, but still no evidence of significant difference.

Finally, start and ending trip locations present some interesting patterns. Male and female riders start travelling from the similar areas, but the destination for males spreads farther apart by comparison. It is unclear why this phenomenon occurs with the information available.  

![trip_duration](images/trip_duration.jpg)

## Preliminary Conclusions

There is not enough evidence to conclude than men are willing to participate in the bike program more so than women based on the data provided. Surveys could serve as a useful tool to complement this study. Since there is no significant difference in trip duration, it would be worthwhile to gather information related to the participant's routine/preferences. For example, female riders may prefer certain areas of the city depending on crime rate. A follow-up assignment would consist on preparing random surveys to gather information about the participant's experience. The data collection system is robust enough and resources should definitely be assign to maintain the program running for the city. 

### Copyright

Arturo Wilson (C) 2020. All Rights Reserved.






## Background

You are well on your way to becoming a programmer and Excel master! In this homework assignment you will use VBA scripting to analyze real stock market data. Depending on your comfort level with VBA, you may choose to challenge yourself with a few of the challenge tasks.

### Before You Begin

1. Create a new repository for this project called `VBA-challenge`. **Do not add this homework to an existing repository**.

2. Clone the new repository to your computer.

3. Inside your local git repository, create a directory for both of the VBA Challenges. Use the folder name to correspond to the challenge: **VBAStocks**.

4. Inside the folder that you just created, add any VBA files. Theses will be the main scripts to run for each analysis.

5. Push the above changes to GitHub or GitLab.

### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - Use this while developing your scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.

### Stock market analyst

![stock Market](Images/stockmarket.jpg)

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

![hard_solution](Images/hard_solution.png)

2. Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

### Other Considerations

* Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller and will allow you to test faster. Your code should run on this file in less than 3-5 minutes.

* Make sure that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.

## Submission

* To submit please upload the following to Github:

  * A screen shot for each year of your results on the Multi Year Stock Data.

  * VBA Scripts as separate files.

* After everything has been saved, create a sharable link and submit that to <https://bootcampspot-v2.com/>.

- - -

### Copyright

Trilogy Education Services © 2019. All Rights Reserved.
