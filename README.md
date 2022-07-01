# VBA of Wall Street
## Green Stocks Module 2 Challenge

## Overview of Project
### For this project I wanted to help Steve get a better understanding of the changes in multiple stocks from the year 2017 to 2018. Steve is already surprisingly good at excel but I wanted to show him the power of VBA and how fast it can be. I wrote a Macro to track the multiple stock volumes and the percentage of return for the given year. Afterword’s I wanted to see if I can make the macro faster for a busy person like Steve.

## Results
### First, Steve wanted to look at the stock "DQ" for his parents at the start. After I found out that the stock was returning -62.6% in the year 2018. I wanted to look at multiple stocks to compare them. I found the "RUN" stock had the highest return out of the twelve stocks I looked at. Which would help Steve get a better idea for his investments for his parents.

### Second, after writing the first macro to run through all the stocks using a "ticker Index", for loops, and if then statements. I put a timer on the function to see how fast (or slow) the macro was running. The macro took 1.117 seconds to run. Knowing time is money I wanted to speed up the process and make it accessible on any year in the worksheet. I also knew I was repeating steps in the macro that I could cut out or simplify.

### Finally, I added a message box so we could manually enter what year we wanted to analyze. I cut down the amount of for loops and the switching between worksheets. I finally got a code that started to run faster. After some more editing I was able to analyze the spreadsheets 2017 and 2018 in record time. The code took 0.156 seconds to run the spread sheet 2017 and took 0.148 seconds to run the spread sheet 2018. This was amazing time for it has cut down the original run time down by 1/10th of what it used to run. 

## Summary
### Advantages and disadvantages of refactoring code in general
#### Refactoring code is great to simplify the process a macro is running. The problems I was running into was changing one thing and forgetting to change the following code. This made the code give me errors and left me frustrated. After you work out the new errors you are surprised to see the amount of filler you had. I would say a huge advantage of refactoring code is to really understand what we are making this macro for. At first, I looked at the "DQ" stock and provided everything needed such as its volume and return. Taking it a step further I could run multiple stocks, get the same results, and compare the stocks better. I really wanted to see how fast I could run the code so it was interesting looking back over and over to see what I could rearrange or cut out.

### Advantages and disadvantages of refactoring code of the original and refactored VBA script
#### At first, I did not make the best arrays for the Volume and the returning rate. I made arrays to track each ticker on its own. having the code loop first then output the results. This speed up the process tenfold. I also wanted to format the cells at the end since I would already be in the spreadsheet I wanted to format. I had to really investigate the code and read what I was doing. Every change I made seemed to make error. in the end what worked for me.
