# VBA_Challenge
Assignment
Report: Deliverable 2 Requirements


The Aim of this Project


This project deploys a technic of refactoring code to explore ways to edit an existing code designed to analysis data of Stock Market Dataset by adding new functionality. The primary aim is to “make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.” To achieve such objective, this project downloaded an existing VBA starter code and refactor it to deliver an efficient output analysis of a given stock dataset within a very short time. 

Written Analysis of Results


	The first step of refactoring the starter code is creating a tickerIndex and set it equal to zero before looping over the rows.
 
	The above image also demonstrates the creation of Arrays for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

	Below image show how tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays 
 
	Below is the image of a refactored script looping through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices
 

	After running my refactored script the outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module

 
Image: Result of VBA Analysis 2017
 
 Image: Result of VBA Analysis 2018

Summary  

Advantages

	Code Refactoring makes the code more extensible for adding many other functions to it. It also helps in increasing the flexibility of the code and by this the capability of code increases.
	After refactoring, the code is fresher, easier to understand or read, less complex and easier to maintain.

Disadvantages

	It is time consuming at some times and one can get lost on how proceed with the coding process.
	Some time I made mistake it took me a while to find where the mistake was.

Pros

	The main goal of code refactoring is to make it easy to enhance and maintain in the future

Cons

	There is a high possibility of introducing bug into a code when refactoring. For instance, in my effort at refactoring the VBA script for the assignment I encountered a lot of bugs that I must find ways to debug them after getting error 9 messages.
