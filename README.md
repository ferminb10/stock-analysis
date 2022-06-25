# vbastock-analysis
Performing analysis on green stock data to uncover trends



Overview of Project: Explain the purpose of this analysis.
The purpose and background are well defined (2 pt).

Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
The analysis is well described with screenshots and code (4 pt).

Summary: In a summary statement, address the following questions.
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
One of the disadvantages of refractoring code for the particular application of analyzing thousands of stocks is that the program may take a long time to execute. You can easily make this to your advantage if you're well educated in economics and you have tailored the code to a limited stock pool.

There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?


## Overview of Project
This weeks challenge helps us build upon our skills learned in the VBA module. The client wants to do a do research for 
## Analysis and Challenges
The analysis had two tasks. First, a solution code was refractored to loop through all stock data one time in order to collect the same information that you did in this module. Then, it was determined if whether refactoring the code successfully made the VBA script run faster, the results shown below.


![VBA_Challenge_2017](https://user-images.githubusercontent.com/107658895/175760234-a2d674a0-997c-461e-9265-4acbc96ad7cb.png)


The second task was to visualize the percentage of successful, failed, and canceled plays based on the funding goal amount. This was done by honing our excel skill of the 'countif()' function. For our purpose, this function counted the outcomes for the different goal ranges. This result was populated in a table and specifically filtered for the subcategory 'plays' in each goal range by "Number Successful", "Number Failed" and "Number Canceled". These criterias' percentage of successful, failed, and canceled projects were calculated for each goal range. A line chart was generated to visualize the relationship between the goal-amount ranges on the x-axis and the percentage of successful, failed, or canceled projects on the y-axis, as shown below.

![Outcomes_vs_Goals](https://user-images.githubusercontent.com/107658895/174226764-5bfb1740-de0d-49f5-8afe-95b9ca2a5642.png)

The most difficult thing for these tasks was making sure that you were filtering for what was being asked. For me, I had trouble, initially, filtering the correct data from the required 'subcategories' when doing the countif function. I overcame this following the tutorials in the module 1 to get the correct Outcomes based on Goal graphs.

## Results
Based on the results of the Theatre Outcomes by Launch Date, I can say that people really like attending plays come summer time. On average, I believe this is around the time of the year when people can take their vacations and do more leisurely activity. There also seemed to be a spike in failed outcomes around September-November which coincided when there were no canceled theater outcomes. For now, the best chances to have successful theater outcome is on February to September.

It seems that your play's outcome has a better chance of being successful than failing if its goal range is less than 20,000 or from 35,000 to 45000.

The limitations of this dataset is that it only provides a narrow time window. This data isn't as representative as if it had it for a wider time range. It also didn't consider every country. I'd like to see how these charts looked for the top 3 countries that spent the most money. I'd also like to see which countries have donated the most money and for which genre.
