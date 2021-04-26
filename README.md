# stock-analysis
📈 Learning how to use VBA in Excel to analysis and predict stock

## Project Overview
The project’s background was to evaluate green energy company stocks to assess the potential for diversification of investment funds. The goal was to learn VBA and to refactor VBA script for optimization.
 
### Results
The initial VBA script was created and run with the file [green_stocks](https://github.com/RuthLD/stock-analysis/blob/main/green_stocks.xlsm). For the 2017 stock data, the script had a run time of 1.621094 seconds, as seen in Before_Refactor_2017 image. ![Before_Refactor_2017](https://github.com/RuthLD/stock-analysis/blob/main/Resources/Before_Refactor_2017.png)

For 2018 stock data, the script had a run time of 1.292969 seconds, as seen in Before_Refactor_2018 image. ![Before_Refactor_2018](https://github.com/RuthLD/stock-analysis/blob/main/Resources/Before_Refactor_2018.png)

The file [VBA_Challenge](https://github.com/RuthLD/stock-analysis/blob/main/VBA_Challenge.xlsm) was used for to refator the VBA script. After refactoring the script, stock data for 2017 ran in 0.203125 seconds, and stock data for 2018 ran in 0.1953125 seconds. Images VBA_Challenge_2017 and VBA_Challenge_2018 show the reduced run times and that the initial analysis has not changed. 

![VBA_Challenge_2017](https://github.com/RuthLD/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png) ![VBA_Challenge_2018](https://github.com/RuthLD/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

#### Key Changes in the Script
In the initial analysis of all stock by year, the starting and ending price were initialized as variables.

* _Dim startingPrice As Single_
* _Dim endingPrice As Single_
 
The following three changes made when refactoring the script influenced the run time of the script for efficiency. 

1) The creation of a tickerIndex variable.
* '1a)Create a ticker Index
	* _tickerIndex = 0_
 
2) The creation of output arrays for tickerVolumes, tickerStartingPrices, and tickerEnding Prices.
* '1b) Create three output arrays
	* _Dim tickerVolumes(12) As Long_
	* _Dim tickerStartingPrices(12) As Single_
	* _Dim tickerEndingPrices(12) As Single_
  
3) Ensuring that the arrays’ starting value was set to zero at the beginning of each loop.
* '2a) Create a for loop to initialize the tickerVolumes to zero.
	* _For i = 0 To 11_
	* _tickerVolumes(i) = 0_
	*  _tickerStartingPrices(i) = 0_
	*  _tickerEndingPrices(i) = 0_
	*  _Next i_
 
### Summary
#### Advantages of Refactoring
Refactoring will start with a preexisting outline of the script, and the code can be used with the modules already in place.

#### Disadvantages of Refactoring
A firm understanding of the VBA syntax is needed to make the script more efficient.
 
#### Pros and Cons
When refactoring the original script for this project, the outline provided had all of the formatting information in place. The ticker array was set up beforehand, saving a lot of time. There was a significant improvement in the run time of the script after refactoring.
The syntax’s exact requirements meant that the order I edited the script could return an error before I finished the edit.
