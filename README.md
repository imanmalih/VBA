# VBA-Challenge
To start my VBA code, I focused on the first worksheet and listed and declared all my constants.
Then I assigned them to their appropriate values. 
I continued to assign values and varaibales, as needed, when going through the For loop and conditionals. 
I decided to create my For loop and have it loop through the entire 2018 Worksheet until the final row that held value. 
In order to pick the first opening price and final closing price, per ticker, I had to create a condtional that would pick the first opening price based on where it was running on the ticker. 
It would start at a row with a new ticker and pick up on the opening price.
As it continued to loop, it would pick up if the next row had the same or different ticker label. 
If the ticker would be different, It would grab the closing price for the ticker it is on and then grab the next rows opening price.
I then got the yearly change by subtracting the closing price from the opening price.
From the yearly change, I was able to get a percentage change per ticker. 
I divided the yearly change by the opening price and assigned the value to round two decimal places with the percentage sign in order to clarify the percent change. 
I created another conditional format in order to change the cell colors in the yearly change column. 
In order to change the cell colors, I had to create a conditonal that stated if the cell value was greater than 0 it would change to the green color otherwise do red. 
When it came to the TotalStockVolume, I wrote the code to pick up the value in volume and add all the volume with the same ticker symbol. 


To simplify and clean the data some more, I created a more concise summary table for the data we could make use of. (The greatest percent increase, the greateset percent decrease, and the total volume). 
I created another conditional to run through the percent change column and pick up on the greatest percentage and lowest percentage and store that value and its ticker in the seperate table. 
Finally, I created a last condtional in order to pick the greatest stock volume with its ticker and put this value in the seperate tables Total Volume row. 
In order for this code to run through all three worksheets, I had to declare "ws" as worksheet and add "ws." before all cells/ranges in order for the code to know to repeat this through each worksheet. 
