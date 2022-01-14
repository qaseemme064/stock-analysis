
# Stock Analysis
# Stock analysis with VBA

## **Overview of Project**
Steve wants to research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although our code works well for a dozen of stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.


### **Purpose**
The purpose of this project is to refactor VBA code . Refactoring will run the code to loop through all the data one time in order to collect theinformation. The main purpose for this is to increase the efficiency of the original code. by refactoring

---


## **CODING IN VBA**


First of all I copied the original code that was done in this module. Then starting adding a code for refactoring: 
1)	 First created a ticker index to zero 
2)	Created three output arrays
3)	Created a for loop to initialize the tickerVolumes to zero.
4)	Looped over all the rows in the spreadsheet
5)	Then within the loop increased volume for current ticker
6)	 Checked if the current row is the first row with the selected tickerIndex.
7)	checked if the current row is the last row with the selected ticker
8)	 Increased ticker index
9)	Loop through your arrays to output the Ticker, Total Daily Volume, and Return.

So below is the code that is written in VBA

    1) Create a ticker Index
   tickerindex = 0
    
    2) Create three output arrays
   Dim tickervolumes(12) As Long

   Dim tickerstartingprices(12) As Single

   Dim tickerendingprices(12) As Single
        
    3) Create a for loop to initialize the tickerVolumes to zero.
   For i = 0 To 11

   tickervolumes(i) = 0 

   Next i
   
    
    4) Loop over all the rows in the spreadsheet.
   For i = 2 To RowCount
    
    5) Increase volume for current ticker
tickervolumes(tickerindex) = tickervolumes(tickerindex) + Cells(i, 8).Value 'stores ticker volumes
    
     
        
    6) Check if the current row is the first row with the selected tickerIndex.
   If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1).Value <> tickers(tickerindex) Then

   tickerstartingprices(tickerindex) = Cells(i, 6).Value 'stores ticker starting price
   End If
        
        
    7) check if the current row is the last row with the selected ticker
   If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then

   tickerendingprices(tickerindex) = Cells(i, 6).Value 'stores ticker ending price

   End If
            

    8) Increase the tickerIndex.
If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then

tickerindex = tickerindex + 1

End If
    
Next i
    
    10) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
For i = 0 To 11
        
Worksheets("All Stock Analysis").Activate

Cells(4 + i, 1).Value = tickers(i)

Cells(4 + i, 2).Value = tickervolumes(i)

Cells(4 + i, 3).Value = (tickerendingprices(i) / tickerstartingprices(i)) - 1

Next i

---

## **RESULTS**

After running the code I received the following results for 2017 and 2018:

![2017 output](https://user-images.githubusercontent.com/96033163/149448001-8f2d6e4f-1bdf-465c-9e57-1450f44709f6.jpg) ![2018 output](https://user-images.githubusercontent.com/96033163/149448249-e78009fd-6f4b-482c-b273-ec2efa9fbd41.jpg)

In 2018 only TERP was in negative but in 2018 almost all of them is in negative

By refactoring the code the running time of the code reduced drastically as shown in below comparison

![comparison of stock 2017](https://user-images.githubusercontent.com/96033163/149448445-a3b8dcfc-3d0f-4afb-884e-7eb0d7daaf25.jpg)

![comparison of stock 2018](https://user-images.githubusercontent.com/96033163/149448447-4acb7836-5c6d-4042-9383-6762d75c2462.jpg)

---

## **SUMMARY**

1. **What are the advantages or disadvantages of refactoring code?**
    ### *Advantages*
   * Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

   * Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be our entry point to working with the existing code at a job.

    ### *Disdvantages*
    * The outcomes of the testing may affect by refactoring

    * In the original code the code may scattered in different locations due to which it may take time to refactor the original code
    
    * A very good hold of understanding the syntax of vba to make script more efficient


&nbsp;

2.  **How do these pros and cons apply to refactoring the original VBA script?**
    
    Refactoring is like cleaning up the orderly house. So the more time you spent the more cleaner the code will be. In this challenge most of the stuff was set up before handed. There was a significant improvement in the run time.
    
---