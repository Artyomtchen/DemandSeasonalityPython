# DemandSeasonalityPython
Time series decomposition to get trend and weighted average seasonality. In this example by country and product, where years closer to TODAY() are weighted more in the average
Given the input consisting of country, oil product, year, month, and demand value (in thousand barrels per day), the code
a) calculates the trend of the data
b) caculates seasonality and removes the trend from seasonality
c) removes one lowest and one highest outlier in calculations
d) considers only last 10 years
e) weights more recent datapoitns more than more remote datapoints in the past
f) finally calculates average seasonality for each country and product for each month (1 to 12)
We adapted this code to country and product granularity, but the code can also be used on deeper or less detailed levels with only a few modifications
