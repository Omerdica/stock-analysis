# Overview
### Steve provided me with an Excel files containing information on 12 different stocks. He was asked by his parents to help him choose what would be best investment for them and what returns they would see based on the data provided from the last 2 years. First we looked at the DQ stock. After compiling data we realised it would be beneficial to run the other 11 stocks to see how they compare to the DQ stock. 

# Results
### Refactored Analysis
### After creating DqAnalysis, we realised that Steve's parents would not benefit from investing in this specific stock. We determined it would be best to run data on all stocks. We had to create a new macro that would help us provide more in-depth information for Steve. The first VBA code we made was to break down the total volume for each company and their annual return. 

<img src="C:\Users\rolli\OneDrive\Desktop\Almir School\Green_Stocks\Resources\Original_Code.png">

<img src="C:\Users\rolli\OneDrive\Desktop\Almir School\Green_Stocks\Resources\Original_2017.png">

<img src="C:\Users\rolli\OneDrive\Desktop\Almir School\Green_Stocks\Resources\Original_2018.png">

### The refactored code was created to see if a different approach would provide more specific information for Steve. By doing this, we were able to increase the speed of the code and also clean up the data. We started by adding indexVolumes and creating new For loops.   

<img src="C:\Users\rolli\OneDrive\Desktop\Almir School\Green_Stocks\Resources\Refactored_Code.png">


<img src="C:\Users\rolli\OneDrive\Desktop\Almir School\Green_Stocks\Resources\Refactored_2017.png">

<img src="C:\Users\rolli\OneDrive\Desktop\Almir School\Green_Stocks\Resources\Refactored_2018.png">

### By looking at the the screenshots we can see that the code impreoved significantly. 

# Summary 
### I took the arrays and moved them to the beginning of the macro to help clean up of the code. We created new For loops based on tickerIndex that were provided in the module. By adding the tickerIndex it helped the For loops better understand of what specific data we were looking for. One of the biggest advantages that i have seen through out this whole proccess was that we took a code that was woring and made it perform faster and still got the same results. One disadvantage to refactored code is that when we start changing things that work there is a greater chance of getting errors through out the code. When I first introuduced the tickerIndex this created many issues to the original code. For loops had to be written in a way that we had to state what was a tickerIndex and its place in the For loop. 











