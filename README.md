***# Stock Analysis & Refactoring Code for Improved performance.  ***

### Overview of Project:
This challenge requires practice of VBA utilizing starter code provided in this Challenge to refactor the original script, which looped through the data one time to collect all of the information required for analysis. The key desired outcome is for refactored code to run faster than it did the original script.

Based on the explanations of Trilogy Education Services, refactoring is a key part of the coding process. When refactoring code, developers aren’t adding new functionality;instead the goal is to make code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. 

### Analysis' Purpose:

The purposes of this analysis is to analyze ticker data for years 2017 and 2018.  The data consists of the following information (note: below, the name of each column is italized for visualization of data structure):
1. the stock's symbol or _ticker_ (12), 
2. _date_ when the shares were traded, 
4. the ticker's _open_ price
5. the ticker's  _high_ price
7. the ticker's _low_ price, 
9. the ticker's _close_ price, 
10. the ticker's _adjusted close_ price, 
11. and the ticker's daily _volume_ of shares traded

The Analysis reviews the total daily volume and yearly return for each stock. Daily volume is the total number of shares traded throughout the day, measuring how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year. 

### Results

#### Stock Performance

When running an analysis for all stocks in 2017, the analysis showed that all stocks reviewed returned a favorable return in 2017, except ticker symbol: "TERP".    All favorable returns for 2017 ranged from 5.5% to 199.4%.  

2018 was not as a favorable year as 2017.  Returns for 2018 depreciated for all tickers, except two stocks, as listed below, otherwise it was a sea of red.  
1. "ENDPH", which had 81.9% positive return for the year, and:
2. "Run" which returned a 84% favorable result for the year

#### Code Refactoring Performance

As the images depict below by refactoring the orignal code we greatly improved performance.  For reference the orignal code to run the 2017 analysis took 1.777344 seconds.  By refactoring the code with efficiencies, speed went down to 0.171875 seconds.  That is a reduction of 1.605469 seconds!
![This is an image](https://github.com/Juligi/Stock_Analysis/blob/main/Resources/VBA_Challenge_2017.png)

The same story is true with the 2018 stock analysis completed as shown in the following capture. ![This is an image](https://github.com/Juligi/Stock_Analysis/blob/main/Resources/VBA_Challenge_2018.png) The refactored code ran in 0.1953125 seconds for the year 2018 compared to the original code, which took 1.636719 seconds.  The refactored code shaved aproximately 1.44 seconds off. 

