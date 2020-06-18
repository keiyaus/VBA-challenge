# VBA-challenge

Note on finding Yearly Change in stock price and percentages:

I have defined the Opening and Closing Prices of each ticker to be the one that has a corresponding trading volume (i.e. <> 0 ), just the same as how we would normally define them in a real stock market.

Hence, the loop I created will not automatically return the first row (which is always January 1 without any trading volume as it is a public holiday)and the last row of each ticker category. Instead, it considers the trading volume argument to be sure we find the actual Opening and Closing Prices.

As a result, the numbers shown in my screen shots for each year will look different from the ones in the grading rubric. But I have selectively checked a number of the datapoints in all three worksheets and believe the calculataions for each ticker are correct.