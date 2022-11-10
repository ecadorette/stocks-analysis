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

 I ran into an error 6 'overflow' that I was not able to debug myself. After some (a lot) of help with a tutor we were able to untangle the bits of code that I had confused (turns out arrays counts start at 1 while indexes start at 0). After debugging the code we were able to get the refactored macro running properly. As shown below, the macro with the nested loops ran in .91 seconds and the refactored macro ran in .17 seconds. 
 
Macro with nested loop:
 
 ![AllStock_VBA_Challenge_2018](https://user-images.githubusercontent.com/114450503/200985312-27d1aecd-b0b7-4074-a38e-d37064f47806.png)
 
Refactored macro with one loop:

![Refactored_VBA_Challenge_2018](https://user-images.githubusercontent.com/114450503/200985369-d6c1bd06-e9a6-4b4b-8d4b-a0d9ef12eb96.png)

As far as results of the data are concerned, as we can see below most stock returns plummeted in 2018, with RUN and ENPH as the only exceptions at an 80+% return rate.

![2017 Stock](https://user-images.githubusercontent.com/114450503/199878774-ce88c208-3cc2-40dc-965d-f5c423b6f334.png)

![2018 Stock](https://user-images.githubusercontent.com/114450503/199878794-ffb2a2c9-4cf4-4b1c-bfde-aee42cf56f0b.png)

### Summary

The main advantage to refactoring code like this is to cut down on processing time; the larger the datasets become the more important each second or tenth of a second becomes. It can also cut down on repetitive code and redundancy. In the right hands, condensed or generalized code may become infinitely more flexible and easier to scale as datasets grow over time. One disadvantage to refactoring could be that the complexity of the code increases. Depending on the builder, their organization of the code, and the clarity of their comments (if any)the code can become more difficult to understand and maintain by collaborators or even that same builder down the road.

The above pros and cons directly apply to this specific project as we started by analyzing DQ for a specific year, to reviewing all stocks over two years. We generalized the code and made it faster as well as more flexible to accomodate for the larger dataset. The downside is things can start to get a little confusing if the code isn't organized well and/or the notes are poor.
