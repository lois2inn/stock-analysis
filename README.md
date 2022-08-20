# Stock Analysis - Stock Performance and Runtime Performance of Refactored VBA Script

## Overview
Steve, a financial analyst, wanted to analyze some green energy stocks for years 2017 and 2018. To help Steve, Excel Macros using Visual Basic for Applications (VBA) were created.

### Purpose
The purpose of this project was to compare the stock performance between 2017 and 2018. The original code written to analyze stocks was refactored so that the execution times can be compared. 

## Results

### Comparison of Stock Performance between 2017 and 2018
The yearly returns of 12 different stocks were compared between 2017 and 2018.  

![Stock_Returns_2017_2018.png](/resources/Stock_Returns_2017_2018.png)

- There was a decrease of returns across stocks in 2018 compared to 2017 except for RUN and TERP tickers
- RUN showed a growth of around 80% in 2018 than the previous year.
- Though TERP had below zero returns in 2017 and 2018, the decline slowed in 2018. 
- DQ did extremely well in 2017 with almost 200% yearly returns followed by SEDG.  
- ENPH and FSLR have over 100% returns in 2017.  The growth of ENPH slowed by 47% in 2018.
- Except for RUN, the rest of the stocks did not fare well in 2018 compared to 2017.

### Comparison of Execution Times between Original and Refactored Scripts
The execution time of the original script and refactored script were recorded.  
<table>
   <tr>
    <td><b>Original Script Run time for year 2017</b> </td>
    <td><b>Refactored Script Run time for year 2017 </td>
   </tr>
  <tr>
    <td><img src="/resources/Original_2017.png" width="400" border="5px"/> </td>
    <td><img src="/resources/VBA_Challenge_2017.png" width="400"/> </td>
  </tr>
  <tr>
    <td><b>Original Script Run time for year 2018 </b> </td>
    <td><b>Refactored Script Run time for year 2018 </b> </td>
  </tr>
  <tr>
    <td><img src="/resources/Original_2018.png" width="400" border="5px"/> </td>
    <td><img src="/resources/VBA_Challenge_2018.png" width="400"/> </td>
  </tr>
</table>

- The original Script run time for 2017: 0.5898438 seconds. The refactored Script run time for 2017: 0.1054688 seconds
- The script run time for 2017 with refactoring decreased by 82%
- The original Script run time for 2018: 0.5898438 seconds. The refactored Script run time for 2018: 0.109375 seconds
- The script run time for 2018 with refactoring decreased by 81%
- Refactoring the code proved beneficial as the run time decreased over 80% for this dataset.

## Summary

**Advantages and Disadvantages of Refactoring Code in general**
- Refactoring helps improve internal code by making many small changes but without changing the code's external behavior. It encourages a more in-depth understanding of the code, thereby making the code easier to understand and read. It improves maintainability and makes it easier to spot bugs or make further changes. Refactoring code may improve the performance of an application.
- Imprecise refactoring can introduce new bugs and break the existing functionality. Refactoring, if not planned for,  will take extra time. This can lead to delays and extra work for the developer. Testing refactored code can be cumbersome if test cases are not in place.


**Advantages and Disadvantages of the original and refactored VBA script**




