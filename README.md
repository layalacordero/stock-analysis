# Analyzing stocks with VBA

## Overview of Project
In this project, we will use VBA to sort through all the data in
the worksheet at once with a refactored code of our Module 2
code. We will add to the refactored code and then analyze both
codes to see if the refactored code can run faster than the
code we typed previously.

## Purpose
The purpose of this analysis is to help a client named Steve to
look at a company called Daqo and see the yearly return from their
stocks. By analyzing the returns, we can help Steve explain to his
parents whether investing in the company is a good idea.

## Analysis and Challenges
By using a `tickerIndex` I was able to access the index needed for
four different arrays and with the for loops not being nested
it helped the program run quicker.

### Refactored Code
![refactored-code](https://github.com/layalacordero/stock-analysis/blob/main/Resources/refactored_code.jpg)

### Original Code
![original-code](https://github.com/layalacordero/stock-analysis/blob/main/Resources/original_code.jpg)

As you can see in the images below the refactored code ran faster

### 2017 Refactored Analysis
![2017-refact-code](https://github.com/layalacordero/stock-analysis/blob/main/Resources/2017_refactored_runtime.jpg)
### 2018 Refactored Analysis
![2018-refact-code](https://github.com/layalacordero/stock-analysis/blob/main/Resources/2018_refactored_runtime.jpg)
### 2017 Original Analysis
![2017-original-code](https://github.com/layalacordero/stock-analysis/blob/main/Resources/2017_original_runtime.jpg)
### 2018 Original Analysis
![2018-original-code](https://github.com/layalacordero/stock-analysis/blob/main/Resources/2018_original_runtime.jpg)
## Summary
After looking at the data, refactoring the code as shown in the analysis
has helped the program run faster. By switching the order of the for loops,
it made the code run efficiently and still gives an accurate report of
the worksheet.
