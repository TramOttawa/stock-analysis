# Stock-analysis
## Overview of project
### Background and purpose
* Steve needs help analyzing stocks market to support his parent’s investment decisions. With my knowledge about Visual Basic Application (VBA) I am willing to help him to provide each stocks annual volume and return on investment (ROI).
* He loved being able to analyze each stock at the click of a button and now wants to expand his research beyond the 12 green stocks. He wants to analyze a higher volume of stocks and to save time, a refactoring of VBA coding was used to improve the execution time.
## Results
### Refactoring the Code
* 3 new arrays: -tickerVolumes(12) to hold volume -tickerStartingPrices(12) to hold starting price -tickerEndingPrices(12) to hold ending price
* The above 3 arrays store performance data for each stock when a for-loop runs analysis on them. 
* Matching the 3 performance arrays with the ticker array is done by using a variable called the tickerIndex.
* Nested For-Loops are used to loop through the data and complete the analysis.

#### 2017 vs 2018 Stock Performance
There is a huge change in the 2017 performance of Green Stocks vs 2018. Only 2 of the 12 stocks (ENPH and RUN) produced a positive ROI in both years.

![Stocks_2017](https://user-images.githubusercontent.com/100484606/159107810-71958f2d-7cb6-41a8-a8e2-cbd5684b7b37.JPG)

![Stocks_2018](https://user-images.githubusercontent.com/100484606/159107817-36c1e1c2-3c43-4241-a085-07f5d6b915f6.JPG)

Steve should advise his parents to look for other stocks in other sectors rather than green stocks as the performances of the green stocks was not positive and have risks. 

#### Execution time
* Comparing between the original script and the refactored script, we can see the improvement dramatically with 86%

![Execution time](https://user-images.githubusercontent.com/100484606/159107458-670d397b-1f08-4041-8e2b-6d1a1ed26fa9.PNG)

* Execution time for 2017 - Original
![VBA_Challenge_2017_Original](https://user-images.githubusercontent.com/100484606/159107476-412fa9a3-5494-4011-9f14-0e39cddf634c.PNG)

* Execution time for 2017 - Refactored
![VBA_Challenge_2017_Refactored](https://user-images.githubusercontent.com/100484606/159107486-ade61b40-6937-419e-80bb-c32daade4e3c.PNG)

* Execution time for 2018 - Original
![VBA_Challenge_2018_Original](https://user-images.githubusercontent.com/100484606/159107493-9d5203f2-bb48-42f7-a292-e19131f3b9c1.PNG)

* Execution time for 2018 - Refactored
![VBA_Challenge_2018_Refactored](https://user-images.githubusercontent.com/100484606/159107465-692d0f27-5bdf-4411-8f33-4eb86d910c92.PNG)


## Summary

### Advantages of refactoring code
The obvious advantage of refactoring code is that it makes it more efficient if you get it right. An 82% reduction in execution time can be huge if analyzing thousands of rows of data.

### Disadvantages of refactoring code
A huge risk with refactoring is that your errors may destroy an already working code. It is highly recommended that you save your original code and any changes you make frequently in case you run into any issues. That way you can always go back a step without needing to start completely over. I personally ran into issues during refactoring and found that using the msgBox script, as well as, testing performance outputs individually helped me identify what was driving my errors.
