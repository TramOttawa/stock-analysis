# Stock-analysis
## Overview of project
### Background and purpose
* Steve needs help analyzing stocks market to support his parent’s investment decisions. With my knowledge about Visual Basic Application (VBA), I am willing to help him to provide each stocks annual volume and return on investment (ROI).
* He wants to be able to analyze each stock at the click of a button and now wants to expand his research beyond the 12 green stocks. He also wants to analyze a higher volume of stocks and to save time, a refactoring of VBA coding was used to improve the execution time.
## Results
### Refactoring the Code
There are some main codings were used to analize the stocks as follows:
* 3 new arrays: tickerVolumes(12) to hold volume; tickerStartingPrices(12) to hold starting price and tickerEndingPrices(12) to hold ending price
* The above 3 arrays store performance data for each stock when a for-loop runs analysis on them. 
* Matching the 3 performance arrays with the ticker array by using a variable called the tickerIndex.
* Nested For-Loops are used to loop through the data and complete the analysis.

![Refactored coding 1](https://user-images.githubusercontent.com/100484606/159109147-0e8e9bf4-8cc8-45af-aee4-5643f4c2f40b.PNG)

![Refactored coding 2](https://user-images.githubusercontent.com/100484606/159109161-e2e20085-8640-4889-b287-fc639d822f20.PNG)

#### 2017 vs 2018 Stock Performance
* Green-energy stocks in 2017 had a high ratio of positive yearly returns (only one green-energy stock (TERP)) had a negative yearly return. Analysis from 2018 showed a completely different picture. The majority of stocks had negative returns. The drop was significant. The DQ stock had almost 200% yearly return in 2017, but in 2018 the stock dropped with negative 63%.

* Only 2 of the 12 stocks (ENPH and RUN) produced a positive ROI in 2018. Steve should advise his parents to look for other stocks in other sectors rather than green stocks as the performances of the green stocks was not positive and have high risks. 

![Stocks_2017](https://user-images.githubusercontent.com/100484606/159107810-71958f2d-7cb6-41a8-a8e2-cbd5684b7b37.JPG)

![Stocks_2018](https://user-images.githubusercontent.com/100484606/159107817-36c1e1c2-3c43-4241-a085-07f5d6b915f6.JPG)

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

### Advantages and disadvantes of refactoring code
* After refactoring, the code is fresher, easier to understand or read, less complex and easier to maintain. 
* Disadvantages of Code Refactoring is time consuming: You may have no idea how much time it may take to complete the process. It may also land you into a situation where you have no idea where to go.

### Advantages and disadvantages of the original and refactored VBA scripts:
* The obvious advantage of VBA refactoring code is that it makes more efficient if you get it right. An 86% reduction in execution time can be significant if analyzing thousands of rows of data.
* A huge risk with refactoring is that your errors may destroy an already working code. It is highly recommended that you save your original code and any changes you make frequently in case you run into any issues. It is suggested to use the msgBox script to confirm before deleting any coding.
