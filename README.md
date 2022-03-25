# stock-analysis
## Refactor VBA Code and Measure Performance
### Overview of Project:
---The purpose of this analysis is to expand the dataset used preivously to find trends on DQ. We are now looking at the entire set of symbols over the last few years.---To be able to improve the performance of the script running through the data using refactoring
### Analysis and Challenges:
---Using the previous information we add additonal arrays to represent the `tickerVolume(i)`, `tickerStartingprices(i)` and `tickerEndingprices(i)`. Once we have these values we can put it in a `For loop` combined with `If-then` statements to represent the checking of tickers and increasing the volume of each ticker---One Challenge discovered was the requirement to make sure the VBscript running the command buttons were properly pointing to the worksheet.
### Results:
---The findings are when using the above referenced the Run Analysis button returned results that was roughly 7 times faster for both years as shown in the below ![Stock Analyis Refactored 2017](https://github.com/jobloom79/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png) and ![Stock Analyis Refactored 2017](https://github.com/jobloom79/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png) 
### Summary: 
---In summary we see that working to use arrays to help shorten some of the coding and the computing can reduce output time as well as the amount of code that has to be written.---While the original analysis was for one stock that required `nested for loops` and took .78125 seconds in one instance, the new code with arrays proved to be more superior. This code took .140625 seconds for the same year and selecting 12 different symbols.