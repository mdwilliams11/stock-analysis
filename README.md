# Stock Analysis with VBA

## Overview of Project
This project involved helping Steve to analyze multiple different green energy stocks using multiple VBA macros. Starting with the stock data from 2017 and 2018 we wanted to create a macro that would loop through all of the data and return the total daily volume and the annual return on each of the green energy stocks.

Analysis was conducted within this [Macro-Enabled Excel Workbook](https://raw.githubusercontent.com/mdwilliams11/stock-analysis/main/VBA_Challenge.xlsm)

## Results

Almost all of the green energy stocks analyzed performed well in 2017. Almost none of the green stocks performed well in 2018. The only stock with a consistent, good rate of retun was **ENPH**. It seems the green stocks are fairly volatile 


Below is a screenshot of the returns from 2017.

![VBA Challenge 2017](https://raw.githubusercontent.com/mdwilliams11/stock-analysis/main/resources/VBA_Challenge_2017.png)

Below is a screenshot of the returns from 2018.

![VBA Challenge 2018](https://raw.githubusercontent.com/mdwilliams11/stock-analysis/main/resources/VBA_Challenge_2018.png)


After running the original analysis, I refactored the VBA macro to run more efficiently. The The original macro ran the analysis in 0.5664062 seconds for the year 2017 and 0.5546875 seconds for the year 2018. The included screenshots show the refactored code running almost 7 times more quickly. The code was refactored to only have to loop through each row a single time whereas the original code went through every row once looking for a specific stock ticker.

The refactored VBA code for the stock analysis macro can be found [HERE]((https://raw.githubusercontent.com/mdwilliams11/stock-analysis/main/VBA_Challenge.xlsm)


## Summary

