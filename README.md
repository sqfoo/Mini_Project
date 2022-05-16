The code in this file directory is used to scrape down the information about the US stocks which are under the 52 week high stock list. The features of the stocks which are recorded into an excel file, are the stock name, traded volume on that day and which sectors it belongs to.

The library used here are selenium and xlswriter and the information is from https://www.investing.com. The browser used here is Safari.

First, I would access to the webpage which records down the information about the 52 week high stocks and record down their name. Then, I would search their name on this website and scrap down their traded volume and the sector they belong to.

Version 1 is mimic our action when we search it on this website. However, Version 2 directly obtain the result from the url.

The attached excel file is the list of stock which were 52 week high stock on 13 May 2022.
