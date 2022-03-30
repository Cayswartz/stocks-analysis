# stocks-analysis


# Stock Analaysis with VBA

## Overview of Project
We are assisiting Steve who is helping is parent make smarter and more informaed decisions on how to invest. Steve parent's are passionate about green energy and would like to invest green energy solutions such as hydro, geothermal, bio, wind and more! However, Steve's parents at the moment have decided to just invest in DAQO New Energy Corp, or DQ, because that is where they first met at a Dairy Queen. Steve is taking a closer look at the DQ stock but also would like his parents to diversify their stocks. In order to assit them in diversifying their stock he has pulled together information on other green energy options. 

## Results
In this assignment we were able to successfully refactor the code to run more efficiently. For 2017, the factoring successfully cut down the code run time to .140625 seconds and for 2018 I was able to cut it down to .109375 seconds.

![ScreenShot](https://github.com/Cayswartz/stocks-analysis/blob/61b111cdb46a3efc75e22397809aaaf321a51151/Resources/VBA_Challenge_2017.png)
![ScreenShot](https://github.com/Cayswartz/stocks-analysis/blob/61b111cdb46a3efc75e22397809aaaf321a51151/Resources/VBA_Challenge_2018.png)

I was able to obtain this goal through an overall reduction of loops in the coding. First, instead of nesting the j loop within the i loop we were able to create two unique and individual loops. First we created the i loop to initialize the ticker volume to 0 and then moved on to the j loop which ran through and identified the starting and ending prices. At the end of the j loop we also manually increase the tickerIndex with the below code: 

 If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
tickerIndex = tickerIndex + 1
End If

This code allowed for us to automate the increase of a ticker index without having to create an additional loop to move the tickerindex through the line. 

Ultimatley, this refactoring allowed us to create three individual loops without any nesting which created efficiencies across the code. 
            
            

## Summary 

### Advantages & Disadvantages
The benefits of refactoring VBA is that it helps simplify the code making it easier to understand, more mainatabilty and easier to scale as you continue to work with it. Additionally, refactoring requires that you go back through your code closer ensuring that there are no bugs and potential issues that could cause larger issues in the future.

The primary disadvantages of refactoring code is that at times it can be time consuming. At times it can be frustrating because you are diving into code that is already written and working properly to clean it up but at times through this process you can end up breaking it and then having to find the correct way to repair it and create the same fuctionality again. 

As a whole, refactoring can do create things for your code but make sure to keep in mind the time that it could take to rework it. 

### How do these pros and cons apply to refactoring the original VBA script?
These pros and cons pretty directly correlate with my experience with refactoring the VBA script. Being new to VBA, the refactoring definitely took awhile and some troubleshooting as I ran into different errors as I worked through it. This at times could be frustrating but once I got to the end result I was a decrease in the time it took to run my code so ultimately was worth it.
