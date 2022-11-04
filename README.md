# stocks-analysis
Module 2

## VBA of Wall Street
### Overview of Project
The purpose of this analysis was to review the performance of stocks in 2017 and 2018. We accomplished this by creating macros in VBA to loop through the data, find all the distinct tickers, then calculate and return the volume and return rates. Then we made a separate macro to format the results for easy reading.

### Results
The script was originally run with two `for` loops and a separate macro for formatting. I refactored the macro to only have one loop and include the formatting, this was accomplished by creating an index for the tickers instead of looping through them separately.
```
 '1a) Create a ticker Index
 Dim tickerIndex as Single
 tickerIndex = 0
```

 I ran into an error 6 'overflow' that I was not able to debug myself. Because of this I don't have a comparison of execution times. 
 

As far as results of the data are concerned, as we can see below most stock returns plummeted in 2018, with RUN and ENPH as the only exceptions at an 80+% return rate.

![2017 Stock](https://user-images.githubusercontent.com/114450503/199878774-ce88c208-3cc2-40dc-965d-f5c423b6f334.png)

![2018 Stock](https://user-images.githubusercontent.com/114450503/199878794-ffb2a2c9-4cf4-4b1c-bfde-aee42cf56f0b.png)

### Summary

The main advantage to refactoring code like this is to cut down on processing time; the larger the datasets become the more important each second or tenth of a second becomes. It can also cut down on repetitive code and redundancy. In the right hands, condensed or generalized code may become infinitely more flexible and easier to scale as datasets grow over time. One disadvantage to refactoring could be that the complexity of the code increases. Depending on the builder, their organization of the code, and the clarity of their comments (if any)the code can become more difficult to understand and maintain by collaborators or even that same builder down the road.

The above pros and cons directly apply to this specific project as we started by analyzing DQ for a specific year, to reviewing all stocks over two years. We generalized the code and made it faster as well as more flexible to accomodate for the larger dataset. And I managed to massively confuse myself with slightly more complex coding techniques while trying to do so. 
