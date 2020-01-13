# pricebookUpdate
Every year my family has to update the price book for my Dad's business. The price book consists of 15 Excel workbooks that contain descriptions and prices for every item in his inventory. Typically they add 5% to each item and round up to the nearest 50 cents. In the past every single calculation was done on a calculator and entered by hand. I thought this was ridiculous so I wrote a program to do it automatically.

## Getting Started
To run the program please use Pycharm. Open Pycharm and go to File -> Open. Select the priceBookUpdate-master folder.

## Configuration
There is a global variable called multiplier which is the percentage change for all items in the pricebook. I have it set to 1.05 for a 5% increase.

The pricebook folder contains all the pricebook files. Any .xlsx files in that folder will be updated. Note that cell formatting is important, i.e numbers must be formatted as numbers and text must be formatted as text. This allows my program to ignore all cells with text.

## What I Learned
In the process of writing this program I learned how to use openpyxl which is a library for working with Excel files. I also had to use the os library to access the pricebook folder and select the proper files to work with. I excluded the cover page as well as files that are currently open (they have a "~$" in front of their name).

I also had to write a method to round a value up to the nearest 50 cents. It was kind of tricky to always round up, but I solved it by looking at the tenth place digit and using some logic and math.

It was a fun and simple program to write after I found the correct libraries, and I'm glad I could help my family out. It's amazing that this process now takes less than 2 seconds compared to the several tedious days it took by hand.
