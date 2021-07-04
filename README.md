# Stock-analysis

## Overview of Project:
Steve has graduated from Finance. His parents are passionate about green energy, and they wanted to be the first Steve's clients to create an investment portfolio. He needed to analyze the stock market to make decisions as part of this task.

Steve asked for our support to write a code to review the information easily. We helped him and a first code was written. The code was good, but Steve wanted to improve the execution time to add thousands of new stocks. We helped. 

The following information was found with the refactored routine.


## Results:

Comparing the original code versus the refactored code, we can observe:
1.	On performance terms:
Given that 2017 and 2018 have the same stocks, due to this fact, the same code can be used. If the name of the stocks would be changed, we needed to adapt the script asking to the user for their desired stocks.
2.	On Execution time:
With the adapted code working with arrays, all calculations are done behind the scenes and stored very quickly using less resources. In the end, all information is already there and needs one loop to be extracted. There is one pause between the calculation and the keyed data.

On the other hand, each calculation is done and inserted in the spreadsheet with the previous code, so the memory is busy and delays more. Eleven pauses happened because the memory is used after each stock calculation. (See details in the next matrix and images)

|Year|Original Code|Refactored Code|
|---|---|---|
|2017|58824.68|2.375|
|2018|58862.59|2.394531|

**2017 Execution Time**

![ VBA_Challenge_2017]( https://github.com/JackieCortes/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

**2018 Execution Time**

![ VBA_Challenge_2018]( https://github.com/JackieCortes/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)


## Summary:

Refactoring Code can offer different challenges as per example the following ones:
  -	Refactoring helps analyze the code to understand the design and meditate on improving or adding actions.
  -	Time-consuming, depending on whether people are related or not with the routine’s objectives or how it was done, people can spend different amounts of time understanding the script..
  -	To have a crucial reason for refactoring is essential because there is a risk of breaking the code.

I think for this code was a good idea refactoring as it was improved the execution time and also more comments were added. I understand better the actions when I followed each line to fill the required points.

### Note:
*If you want to see the codes please consult the following files in the Visual Basic Menu Option:

*1.Refactored Code "Sub AllStocksAnalysisRefactored()" in module 2 of [VBA_Challenge.xlsm](https://github.com/JackieCortes/stock-analysis/blob/main/VBA_Challenge.xlsm)*

*2.Original Code "Sub AllStocksAnalysis()" in module 1 of [green_stocks.xlsm](https://github.com/JackieCortes/stock-analysis/blob/main/green_stocks.xlsm)*
