# stock-analysis
Analyzing stock data 
# Overview of Project


## VBA Challenges

The challenging part of the analysis was filling the columns titled: "Ticker", "Total Daily Volume, and "Return", with accurate data by using conditional If expressions within nested For loops. After running into overflow bugs and subscript out of range errors, using the local variables window during debugging truly helped correct the project. Another goal of this project was to look at how to improved the efficency of the code using refactoring to optimize the computing power needed to complete the macros. For this we used our initial version of the code which was functional but required the program to complete an itteration of the entire data set for each stock ticker we were looking to analyze. The goal was to develop a refactored version of the original code that required the program to only complete on itteration of the data set and obtain the same data as the first version

## Results
###
To start we developed functional code that allowed for the collection of the total annual stock volume and the year over year performance of each stock we were looking to analyze. The code  provided has the  data that we collected for each stock by completing an itteration of the dataset and inserting the data into the Excel worksheet before moving to the next stock.



## Refactored Code

Through refactoring of the original working code, a more efficient use of computing can be accomplished by reducting the number of total itterations of the data set which results an increase in the speed of the execution of the task. To refactor this code, two components were need to be added to the macro being developed. The first was an index for each ticker that would be itterated for each line of data in the datasets being analyzed. This ment for each line of of data within the dataset the program would identify which ticker was present and store the relevent data related to the index value. The second was a collection of data arrays to store the multiple data points for each stock ticker. As each value saved in the array could be retrieved based on the order that they were collected it was possible to link this data back to the ticker index used. Using these two tools allowed the refactoring of the code




## Summary
## Summary

### Advantages and Disadvantages of Refactoring Code

The process of refactoring code has some advantages and disadvantages for its use. lets look at some of the advantages

 Reduce code smell which is the inculsion of unnessessary, inefficent or unclear coding that makes it harder to read and reuse.  Some examples of evidence of code smell include duplicate code, long method, and long parameter lists.
     - Increases the efficency of the program requiring less computing power to complete tasks
          - Can be a significant factor when dealing with large datasets
 

          
Some of the disadvantages of the use of refactoring of code.
 Requires added time and expense to the development of the code.
     -Increases the complexity of the code and can make is more difficult read and interprete
          - Requires additional variable to be introduced and in some cases additional tasks to make the code functional
          - Requires increased clarity of the notes made in the code to allow other to understand the reason for the use of some code

          
### Pros and Cons of Using Refactoring in for this Code

In the example we showed here there were some pros and cons to the refactoring that was completed to improve the efficency of the code.

 Some of the positive thinks that were the result of the change of the code are as follows:

     - Reduced the time used to complete the analysis by reducting the number of itterations completed to gather the data.  
     
     - Utilized arrays to store data which can be used for other calculations or analysis if a more indepth analysis of the data were needed
     
 Some of the negative factors for the use of refactoring in this code include the following

     - Resulted in a longer code to write. This was due to additional loops that were needed to be added for the use of arrays to store the data as well as increased number of variables added to the task.
     - More complicated to write the code as it required some deeper understanding of coding and requires clearer notes to follow the logic.  This is mainly due the use of arrays to store multiple variables related to one factor being analyzed.  