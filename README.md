# Stock Analysis - Deliverable 2

## Overview
  The goal was to optimize a script pulling data related to stock indexes, and running a cursory gain/loss indicator. This method saved a lot of time ultimately, as it does not run empty rows. 
  
## Results
  For [2017](https://github.com/qklm/stock-analysis/blob/main/VBA_Challenge2017.png), the updated macro shaved off the majority of the time taken for the original macro, coming in at .10s while the other method took .68s. 
[2018](https://github.com/qklm/stock-analysis/blob/main/VBA_Challenge2018.png) had an even larger differential. The original script took .68s again, while the updated macro shaved that down to .09s. In this case the results are not vastly different in  actual output time, but in a data set 10x or 10,000x the size, this is a critical optimization.

## Summary
  The new script has a lot of hardcoded numbers, and that likely is not ideal moving forward. It does however find a significantly faster way to compute the data than a traditional nested loop. As for refactoring: I had a whale of a time getting this done (I think I've seen every error code Excel has to offer). Ultimately simplicity won out, and the team of Learning Assistants got me across home plate. The biggest issue lied in the If Then loops with Error Code 9. Apart from the personal frustrations in this challenge, the refactoring of the code shows the value in diversity of approaches, and in a larger scale would save A LOT of time. 
