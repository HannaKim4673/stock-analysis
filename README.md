# Analyzing Stock with VBA and Excel

## Project Overview

### Purpose
The overall purpose of this project was to help a recent college graduate analyze green energy stocks using VBA for his parents, who are interested in investing in green energy companies. His parents were investing in a green energy company called DAQO, but they wanted to know if they were investing in the best company. While the original code was successful in analyzing the ticker volumes and returns for the twelve companies in the dataset, the student decided to expand his analysis to the entire stock market. Unfortunately, the code did not run particularly fast for just twelve companies, so expanding the dataset to the entire stock market most likely would result in either the code not working anymore or outputting incredibly slowly. As a result, the VBA script needed to be refactored to be able to handle larger datasets and output faster and more smoothly.

## Analysis and Refactoring Results

### 2017 and 2018 Stock Performance Analysis
Looking at the results, the companies indicated by the tickers AY, CSIQ, FSLR, JKS, and SPWR dataset experienced a decrease in volume and return percentage between the years of 2017 and 2018, as seen in the two tables below. Interestingly, the companies DQ, ENPH, HASI, SEDG, and VSLR experienced an increase in volume and a decrease in return percentage between the years of 2017 and 2018 (see tables below). Meanwhile, the remaining two companies RUN and TERP experienced an increase in volume and return percentage between the years of 2017 and 2018 (see tables below).

![](https://github.com/HannaKim4673/stock-analysis/blob/01d31db054c6b2be1574fdb70c44af9352f84591/Resources/2017%20Stock.png) 
![](https://github.com/HannaKim4673/stock-analysis/blob/01d31db054c6b2be1574fdb70c44af9352f84591/Resources/2018%20Stock.png)

### Refactoring VBA Code Results
Overall, the VBA script was successfully refactored. It seems that changing the volume, starting price, and ending price variables in the original VBA script into arrays helped to make the refactored script run faster when executed. To clarify, this is how those variables were defined in the original script:

![](https://github.com/HannaKim4673/stock-analysis/blob/main/Original%20Code.png)

And this is how those variables were defined as arrays in the refactored script:

![](https://github.com/HannaKim4673/stock-analysis/blob/main/Refactored%20Code.png)

For time comparison, the below two screenshots show the run times for the 2017 and 2018 stock analyses using the original script:

![](https://github.com/HannaKim4673/stock-analysis/blob/main/Resources/2017%20Before%20Refactoring.png)
![](https://github.com/HannaKim4673/stock-analysis/blob/main/Resources/2018%20Before%20Refactoring.png)

And the below two screenshots show the run times for the analyses using the refactored script:

![](https://github.com/HannaKim4673/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![](https://github.com/HannaKim4673/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

In general, the refactored script analyzed the stocks about 0.5 seconds faster than the original script.

## Summary

### Advantages and Disadvantages of Refactoring Code
Starting with the advantages, refactoring code generally makes the code more readable for users. Furthermore, the code becomes generalizable in the sense that it no longer refers to a specific dataset and thus can be used with other datasets for a similar analytic purpose. Also, refactored code generally runs more smoothly and quickly in the end than the original. In addition, a programmer may actually improve the code while refactoring by discovering easier or better ways to have the code run while editing it. 
As for disadvantages, a key one is that refactoring can take a long time. Also, it is likely that minor obstacles or errors will present themselves when a code is being refactored, as a result of slight syntax errors and the like, and make the refactoring process take longer than it would otherwise. 

### How those Pros and Cons Relate to this VBA Script
Compared to the original VBA script, the refactored one definitely ran faster and more smoothly, as seen in the message box screenshots in the "Refactoring VBA Code Results" section. Also, the refactored script definitely became generalizable, since it can now be used to smoothly and quickly analyze more stocks, as opposed to just the twelve companies that are included in the [VBA_Challenge.xlsm](https://github.com/HannaKim4673/stock-analysis/blob/main/VBA_Challenge.xlsm) datasets.
In addition to those pros, a huge con that I experienced while refactoring this VBA script was that I made a few small syntax errors that made the refactoring process more frustrating and time consuming than it otherwise would be. For example, one reason why I ended up taking longer than I expected was that one of my variables was misspelled, and it took me a while to figure out why I was getting an error. 
