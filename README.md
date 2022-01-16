# VBA_Challenge

## Overview of Project

Refactoring the module 2 original code in order to run the VBA codes faster. Refactoring helps to remove excessive codes and repition while occupying less space to deliver a better result. The original code used a [for](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/fornext-statement) loop to loop through the tickers, whereas it is also possible to introduce a set of arrays to loop the tickers.

## Results

The Results of the total daily volumes and return of all the 2017 stocks are shown in the image below.[![dasd.jpg](https://i.postimg.cc/v8XZdgtC/dasd.jpg)](https://postimg.cc/9zRhdfSB)

The Results of the total daily volumes and return of all the 2018 stocks are shown in the image below.[![mkd.jpg](https://i.postimg.cc/268gd6LB/mkd.jpg)](https://postimg.cc/CZXvSwNF)

It takes 0.328 and 0.335 seconds to run the 2017 and 2018 with the refactored coding style.[![2017.jpg](https://i.postimg.cc/kM2p7LQ9/2017.jpg)](https://postimg.cc/RqzRRspb)
[![2018.jpg](https://i.postimg.cc/MpLcJMqz/2018.jpg)](https://postimg.cc/yDF6FNDG)

It takes the original code 2.11 and 2.23 seconds to run the 2017 and 2018 by using an additional for loop to go over each ticker.
[![sa.jpg](https://i.postimg.cc/MHX7YkYD/sa.jpg)](https://postimg.cc/T5M5R4FL)
[![dasxz.jpg](https://i.postimg.cc/Fzcj6y2d/dasxz.jpg)](https://postimg.cc/y31ggZ6s)

## Summary

1. What are the advantages or disadvantages of refactoring code? The refactored code is easier to understand, and It takes less space and time to run the code. However, it is very time consuming and impossible to know how long it takes to refine the code.

2. How do these pros and cons apply to refactoring the original VBA script? It makes the code simpler to go over with fewer loops, and it runs the code around 10 times faster.
However, finding a way to refactor the code wouldn't be easy to figure out without the challenge guidelines.

