# VBA Challenge - Stock Analysis 

# **Overview of Project**
The purpose of this challenge was to refactor the VBA script that was previously written to loop through stock data for a particular year, and output calculations including the Total Daily Voume and Return for each ticker in the data set. By refactoring the code, the script is made to be more flexible so it can pull data for any given year that is input, and it can loop through several stocks as they are stored in an array. This saves time as it is much quicker than having to change the code to match the ticker and year that we are looking to pull data for. 

## **Results**

The timer results showed the VBA script ran more efficiently while running the refactored code. Instead of "hard-coding" in the year or ticker that we want to analyze, a more flexible script was created using an array to store ticker data, and by using a new variable to iterate through the arrays.

Prior to refactoring, run times were approximately ~1 second to obtain stock analysis data for each year (2017 and 2018). The timer results for running the refactored code on the same data for these years shows run times of only fractions of a second: 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/90593897/136094440-7251e909-a463-4912-9af6-06dc536aa99e.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/90593897/136094486-29944754-08cb-451b-9f8d-8fa84a43d0bc.png)


## **Summary**

In summary, it was found that the refactored code ran more efficiently than the original VBA script. Refactored, more flexible code would be great for analyzing large datasets as the decreased run time would largely contribute to efficiency as the dataset size is increased.
A disadvantage to refactored code is that it may take longer to develop and the difference in run-time may not be as meaningful or necessary with smaller datasets. 

For example, the refactored VBA script that was developed to analyze stock data would be preferrable with a much larger dataset that holds information on thousands of stocks. A runtime decreased by ~66% ( ~1 second runtime vs .2 second runtime) was already observed with a small dataset, so this saved time could quickly add up to minutes, hours, and/or days with much larger datasets. 
A disadvantage would be that the refactored code took longer to develop, and it may be possible that information is only needed on one stock. In this case, it would be more efficient to use the original script to pull data for a specific ticker and year. 

Overall, it was found that refactored code can be very useful in saving time when running analyses or solving problems. It is important to think about what information is needed from the data, consider the dataset size, and then one can make the best decisicion on how to code to solve the problem. 

