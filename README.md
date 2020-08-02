# stock-analysis
Module 2 - Wall St analysis to help Steve analyze stocks



# **VBA Wall Street Analysis**


## **Assignment Purpose**

The purpose of the assignment is to make VBA code that we had more efficient and run faster.


## **Process**
We start with the existing code and reuse it, but make process a bit more efficient. We achieve that by modifying a few criterias:

### ***1. Year***
We automate the process by establishing year using function 'yearValue'. In this case we are able to run analysis for any year that has data in the existing workbook rather than run a script for each year:

```
Worksheets(yearValue).Activate
```


### ***2. Input Box***
Once we streamlined the process by enabling our macro to run for multiple years, we make sure that any end user can request analysis for any requested/available year. We do that by enabling a pop-up box that allows us to input the year we want to analyze.
https://github.com/jojobear2020/stock-analysis/blob/master/Resources/VBA_Challenge_INPUT_Box_Pop-up.PNG


### ***3. Loops***
To increase efficiency of the process, we use fewer loops to essentially get the same output as in original script. This also means our analysis request may run faster (depending on system capacities).


### ***4. Formatting***
We use conditional formatting to help any end user visualize the final output. 

https://github.com/jojobear2020/stock-analysis/blob/master/Resources/VBA_Challenge_2017%20Output.PNG
https://github.com/jojobear2020/stock-analysis/blob/master/Resources/VBA_Challenge_2018%20Output.PNG


### ***5. Macro Button***
We allow any end user to run analysis by simply clicking on active macro button that already has the complete macro/vba script linked to it. It also allows to see how long it took to run the analysis.
https://github.com/jojobear2020/stock-analysis/blob/master/Resources/Macro_Button_Run%20Analysis.PNG


## **Results**
Even though we modified the script, we are able to see result as expected per assignment. Per given data, 2017 shows that we had better return vs. what we see in 2018. If we had more than two years of data, we would probably have even a better understanding of the trends for all stocks. The main advantage of the modified script is that it allows us to analyze multiple years by using one script and even add more years of data, if needed. 
