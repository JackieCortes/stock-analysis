# stock-analysis

##Overview of Project:

##Results:

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


##Summary:
