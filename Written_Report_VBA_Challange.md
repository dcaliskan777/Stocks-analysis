# Refactoring A VBA Script With An Example Of Stock Analysis 

## Overview of Project

In this project a stock analysis was made by using two vba scripts which are created for a data set which has a dozon of stocks:
Fisrt script, called in this work as "yearValueAnalysis", is the one which is first comes to mind, it is easy to think but takes longer time to execute. Second script, called in this work as "yearValueAnalysisRefactored" is smarter, more efficient but it requires deaper thinking. Sometimes it is sufficient to solve the problem but sometimes the problem requirs to solve it efficiently. In this scenario, there was just a dozon of stocks, so it might be analysed in any way. But, if there would be thousands stocks, an efficient way should be used.  

Accomplishing a task more sufficiently in coding is called **refactoring** which is crucial issue especially in data analysis. Refactoring does not mean adding a new functionality to the script; it is redesigning the script in such a way that it has fewer steps or it uses less memory or it has improved logic to help users to read more easily. These are three main factors of efficiency of a script. In this work, the script "yearValueAnalysis" was refactored in two ways: fewer steps and improved logic, so that "yearValueAnalysisRefactored" was created.

 In the report; the macro "yearValueAnalysisRefactored" is analyzed, challenges and excitements are shortly narrated. The improved efficiency of the subroutine, in terms of elapsed running time,is shown by using pictures on which elapsed time of both subroutines are displayed, for bothe yeras. And finally, advantages and disadvantages of refactoring are discussed in a summary.

### Purpose

The purpose of the project is to examine how the refactoring is crucial in coding.


## Analysis Of The Script "yearValueAnalysisRefactored", Challenges And Excitements

### Analysis Of The Script "yearValueAnalysisRefactored"

The main body of the script consists of only one for loop and only one simple conditional. Simple conditional means here, the structure of "If ...Then ...Else...End If", no ElseIf statement. But, former script contains a nested for loops and several conditionals with ElseIf satatements; these two were origin of the problem. I examined them and pointed out that although an individual ticker is contained around 251 rows in average, for the analysis of stock corresponded to that ticker program is running entirely, to rearch in 3012 rows; and since first row of any ticker is just after last row of the previous ticker, we can retain them both simultaneously and used them seperately.This is the main reasoning in the macro. Therefore I came up with using only one for loop and the following conditional:

> If Cells(j, 1).Value = currentTicker Then
>
>  currentTickerTotalVolume = currentTickerTotalVolume + Cells(j, 8).Value
>
>   Else
>
>  currentTickerEndingPrice = Cells(j - 1, 6).Value
>
>  Worksheets("All Stocks Analysis WRF").Activate
>
>  Cells(tickerIndex, 1).Value = currentTicker
>
> Cells(tickerIndex, 2).Value = currentTickerTotalVolume
>
> Cells(tickerIndex, 3).Value = currentTickerEndingPrice / currentTickerStartingPrice - 1
>
> Sheets(yearValue).Activate
>
> currentTicker = Cells(j, 1).Value
>
> currentTickerStartingPrice = Cells(j, 6)
>
> currentTickerTotalVolume = Cells(j, 8)
>
> tickerIndex = tickerIndex + 1
>
> End If

This is the central portion of the script. If the statement "Cells(j, 1).Value = currentTicker" is False then j'th row is the first row of the next ticker and j-1'th row is the last row of the previous one. This allows us to identify all olutcomes that we need, import them in correct cells and initilize the variables correctly for the next ticker. Notes that the order of sentences after "Else" is important.

This reasoning might create a problem in the last row (last row of the last ticket). In order to avoid this problem I decided to add a fake row at the end of the sheet of the selected year as

> Sheets(yearValue).Activate
> 
> RowCount = Cells(Rows.Count, "A").End(xlUp).Row> 
> 
> Cells(RowCount + 1, 1).Value = "Dursun"
> 
> Cells(RowCount + 1, 6).Value = 1
> 
> Cells(RowCount + 1, 8).Value = 1

In order to keep originality of the sheet of the year, the following sentences are added to script:

> Sheets(yearValue).Activate
> 
> Cells(RowCount + 1, 1).Value = ""
> 
> Cells(RowCount + 1, 6).Value = ""
> 
> Cells(RowCount + 1, 8).Value = ""
 
### Challenges And Excitements

First I conseidered the weakness of the former macro, it was the nested for-loops. It was written in the step by step explanation of the project, we might use tickerIndex for indexing tickers. It was important clue for me to use just one for-loop. I might use tickerIndex for the function of outher loop in the former macro. But, in the steps of the project, two for loops were talked about; it made me quite confused. I was still thinking that there would be a way to decrease number of steps in the script by using just one for loop. I thought a lot, may be 2 hours, without doing anything but planning the code in my mind with just one for loop. At the end I found the way, I was happy with a quite worry, I ran the subroutine, "Run-time error '9'" appered in the window. I was disappointed. I had no idea what kind of error this is. Would it be a logical mistake? I think, the most difficult part in a programming is to encounter a logical mistake, it is very difficult to find out the logical mistake sometimes; sometimes it is needed to change completly the paradighm. I googled and find out the following:

![Run-time Error '9'](https://user-images.githubusercontent.com/99373486/155890334-9af6f0d5-3a15-4afa-ba87-e883383c5ddb.png)

## Results

![](./resources/Outcomes_vs_Goals.png)

### Discussions Of Efficiency Of Refactored Script With The Year 2017

### Discussions Of Efficiency Of Refactored Script With The Year 2018



## Summary


