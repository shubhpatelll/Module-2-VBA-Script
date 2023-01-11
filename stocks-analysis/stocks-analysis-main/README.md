# Stocks-analysis

## Overview of Project
The main goal of this project is to analyze 12 different green energy stocks in the years 2017 and 2018, and identity the stocks with better returns. As an extension of this analysis, original script was refactored so that the client can run the analysis for many stocks for any year in a faster manner.  

## Results
Results are presented in two steps. First the returns of each stock are considered and then the computation time is considered. 
### Best Return
Twelve different green energy stocks in 2017 and 2018 were considered for this analysis. Their tickers are "AY", "CSIQ", "DQ", "ENPH", "FSLR", "HASI", "JKS", "RUN", "SEDG", SPWR", "TERP" and "VSLR". For each year, the volume and the percentage of the return were computed on each ticker. The total volume is reasonably similar for each stock. Therefore, we focus on the return of each stock to identify where to make investments.

The following table shows the "Ticker", "Total Daily Volume" and the return for each ticker in the year 2017.Positive returns are highlighted in green, and the negative returns are highlighted in red. According to the results, the ticker "DQ" provided the best return in 2017.

![AllStocks_2017](https://user-images.githubusercontent.com/112113327/191567387-f70aef7f-ae75-434a-98ae-8b110e012df7.png)

The following image shows the results for the year 2018. According to the results, only stocks "ENPH" and "RUN" provided positive returns. 

![AllStocks_2018](https://user-images.githubusercontent.com/112113327/191567471-9936f24c-bba1-41c2-99ec-9e2833933439.png)


### Refactoring the script
At the beginning of the analysis, the VBA script was written using a nested *for loop*. However, the client requested that he would like run this analysis for the entire stock market. Therefore, the VBA script was refactored using arrays so that nested *for loop* was omitted. The following two images show the part of the script with original nested *for loop* and the refactored script with arrays.
#### All stocks Analysis code: Original scripts with the nested for loop
<img width="495" alt="NestedForLoop" src="https://user-images.githubusercontent.com/112113327/191571929-9d199608-fe72-4fad-80d1-51d370199fdf.png">  

#### All stocks Analysis code: Refactored Script with Arrays
<img width="664" alt="Refactored ForLoop" src="https://user-images.githubusercontent.com/112113327/191571962-11b7410b-c460-4983-ab8c-4bb399f066c3.png">



Now the code run faster than the original script. For example, the run time for 12 stocks in 2018 was 0.293 seconds with the original script, but the run time with the refactored script is 0.078 seconds which is almost 4 times faster. Also, it can be observed  the same efficiency in 2017. For both 2018 and 2017, the run time comparisons are shown bellow.


#### Run time 2018: Original scripts VS Refactored Script
![VBA_Module_2018](https://user-images.githubusercontent.com/112113327/191569604-8f8a8f5f-9804-4589-b646-8faf4b75b8d1.png)   ![VBA_Challenge_2018](https://user-images.githubusercontent.com/112113327/191568984-09d332d0-d1e9-4503-b6ce-06df931ab98d.png)    
  

#### Run time 2017: Original scripts VS Refactored Script
![VBA_Module_2017](https://user-images.githubusercontent.com/112113327/191570099-b3757f76-2b70-4e2d-9b77-258bc65e2ab7.png)     ![VBA_Challenge_2017](https://user-images.githubusercontent.com/112113327/191570147-14ff3c07-a445-407d-acb6-207f49f3b32b.png)



## Summary
In general, refactoring code is an efficient way to reduce run time as well as to the usage of memory in computer programs. However, that is not straight forward for at least for beginners.  
According to the result shown above section, it can be observed that refactored VBA script runs **FOUR times** faster than the original script. There were *for loops* in side another for loop in the original script and it took more computation time, but eliminating inside *for loop* and introducing arrays made the code more efficient. The straight forward method for beginers is the original method. Eventhough, the refactoring is efficient, new programers may take some time to write effecient code directly. 


