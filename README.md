# stock-analysis

## Overview of Project

  Steve's parents made the decision to put all of their money into shares of the Daqo New Energy Corporation (ticker symbol is DQ). Steve wanted to look into Daqo stock for his parents and also, to analyze a handful of green energy stocks in addition to Daqo, so he has created an Excel file containing the stock data he wants us to analyze. 
  In order for Steve, to see the information requested about DQ stock, we first wrote the code to determine its return for the year 2018 and when we run it we could see the table like this:

<img width="355" alt="DQ analysis Excel" src="https://user-images.githubusercontent.com/105990653/175432003-10545800-59a6-4e37-9588-fc75dc628f1d.png">

Then, we wrote the code to perform the same analysis on all the stocks we have in the spreadsheet and the code to select the year to do so. Finally, we wrote the code for a button that would run our code, so at the click of a button, we can analyze an entire dataset like on a picture below:

<img width="517" alt="All stocks Excel" src="https://user-images.githubusercontent.com/105990653/175433059-5fb46268-d146-4765-898a-409b71a92b8b.png">

Now Steve wants to expand the dataset to include the entire stock market over the last few years. Our code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute. 
So we've been given a solution code which we need edit, or refactor, so we could loop through all the data one time in order to collect the same information that we did in this module. 

## Results 

I analyzed data of the results after refactoring of the provided code and making sure it actually works. 
In 2017 all stocks except one (TERP) had a positive return. The stock with the highest return (199.4%) was DQ, which we have been analyzing in the beginning. The total daily volume for DQ in 2017 was 35,796,200 and the total daily volume for all of the analyzed stocks was 3,166,639,100 as you can see on a picture below:

<img width="270" alt="2017 Analysis" src="https://user-images.githubusercontent.com/105990653/175436831-3cd1fea2-5fe7-4b74-a88e-9c4a50c7f915.png">

In 2018 I could see another scensrio for all the stocks. Except for 2, almost all of them had a negative return. The ticker DQ that we were first examining has a return of -62.6 percent and a total daily volume of 107,873,900. Stock with ticker TERP which had the biggest return in 2017 has return of 5 percents and total daily volume is 151,434,700 in 2018. The total daily volume for all of the analyzed stocks in 2018 was 3,306,038,200 as you can see on a picture below:

<img width="264" alt="2018 Analysis" src="https://user-images.githubusercontent.com/105990653/175663704-9b9e23fb-fa58-4464-9168-c957f197e1fa.png">

The following is what I observed regarding the execution times of the original and refactored codes: in 2017, the original code executes in roughly 0.28 seconds, while the refactored code executes in only 0.06 seconds, which is about 0.22 seconds less.
### Original code: 

<img width="258" alt="2017" src="https://user-images.githubusercontent.com/105990653/175676641-de78431d-a606-4988-868d-e1050d8f437c.png">

### Refactored code: 

<img width="281" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/105990653/175665216-127bcf21-6a56-45ad-a16d-97cac7b6d86f.png">

In 2018 was almost same scenario:

### Original code

<img width="261" alt="2018" src="https://user-images.githubusercontent.com/105990653/175665620-11e99cf0-fe70-4539-9933-28632e86e2c7.png">

### Refactored code

<img width="266" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/105990653/175665697-39080915-4129-4930-8cb7-15d3e1e1a6d8.png">

## Summary

1.The original code has a nested loop, which is, in my opinion, can slow down the process of execution. Refactored code doesn't have a nested loop, but 
it took more failure to rewrite it.

2. Comparing the original and refactored code I can see that refactoring code helps to speed up execution, but working on refactoring may require more 
time and resources if code, for example, too long. 










