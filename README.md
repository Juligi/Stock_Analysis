*** Stock Analysis & Refactoring Code for Improved performance.  ***

### Overview of the Refactoring Challenge:
This challenge requires the practice of VBA utilizing the starter code provided in this Challenge to refactor the original script, which looped through the data one time to collect all of the information required for analysis. The key desired outcome is for refactored code to run faster than it did in the original script.

Based on the explanations of Trilogy Education Services, refactoring is a key part of the coding process. When refactoring code, developers aren’t adding new functionality; instead the goal is to make code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. 

### Analysis' Purpose:

The purpose of this analysis is to analyze ticker data for the years 2017 and 2018.  The data consists of the following information (note: below, the name of each column is italicized for visualization of the data structure):
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

When running an analysis for all stocks in 2017, the analysis showed that all stocks reviewed returned a favorable return in 2017, except the ticker symbol: "TERP".    All favorable returns for 2017 ranged from 5.5% to 199.4%.  

2018 was not as a favorable year as 2017.  Returns for 2018 depreciated for all tickers, except two stocks, as listed below, otherwise it was a sea of red.  
1. "ENDPH", which had an 81.9% positive return for the year, and:
2. "Run" which returned an 84% favorable result for the year

#### Code Refactoring Performance

As the images depict below by refactoring the original code we greatly improved performance.  For reference, the original code to run the 2017 analysis took 1.777344 seconds.  By refactoring the code with efficiencies, the speed went down to 0.171875 seconds.  That is a reduction of 1.605469 seconds!
![This is an image](https://github.com/Juligi/Stock_Analysis/blob/main/Resources/VBA_Challenge_2017.png)

The same story is true with the 2018 stock analysis completed as shown in the following capture. ![This is an image](https://github.com/Juligi/Stock_Analysis/blob/main/Resources/VBA_Challenge_2018.png) The refactored code ran in 0.1953125 seconds for the year 2018 compared to the original code, which took 1.636719 seconds.  The refactored code shaved approximately 1.44 seconds off. 

### Summary

The original code got the job done, but by improving and changing the existing code, we were actively able to make the code run more efficiently and faster.  The new code adeptly reran the stock analysis for the 12 tickers reviewed for 2017 and 2018, allowing for quick comparisons.     

#### Advantages of Refactoring Code

According to [Wikipedia](https://en.wikipedia.org/wiki/Code_refactoring#:~:text=In%20computer%20programming%20and%20software,without%20changing%20its%20external%20behavior.), in general, there are 2 key benefits to the activity of refactoring:
1. Maintainabily -- it is easier to maintain and fix issues and bugs because the source code is easy to read and the original intent easier to grasp.  
2. Extensibility -- it allows for extending the capabilities of the source code and the application if it uses recognizable design patterns, providing flexibility where possibly none existed originally.  

"Refactoring improves the design of software, makes software easier to understand, helps us find bugs, and also helps in executing the program faster. There is an additional benefit of refactoring. It changes the way a developer thinks about the implementation when not refactoring" ( [Source](https://aip.scitation.org/doi/abs/10.1063/1.3516393?journalCode=apc#:~:text=Refactoring%20improves%20the%20design%20of,the%20implementation%20when%20not%20refactoring.)).

#### Disadvantages of Refactoring Code

On the opposite, refactoring code can have its challenges, including:
1. Time-Consuming -- it's uncertain how much time it will take to complete the process and where to start.
2.  Chance of Mistakes -- a bigger probability of changing the code to a point your design no longer runs as it did before undertaking the activity of refactoring the code.  

Nonetheless, the activity of refactoring is a wise one and a best practice utilized by programmers to improve their productivity, their know-how, and most importantly, continue the performance optimization of their code and programs created.  

