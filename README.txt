To use this script:
Open the workbook of stocks data you want to analyse.
Go to Developer -> Visual Basic to open the Visual Basic IDE.
In this new window, go to File -> Import File, and select "vba_challenge.bas".
Press run to execute the stocks analyser on all sheets within the workbook.

The expected format is a data positioned at A1 on each sheet with headers, containing columns as follows:
	ticker, date, open, high, low, close, vol

Using a different data structure, or failing to position the table starting at A1, will cause errors.

The files '2018 Analysis.jpg', '2019 Analysis.jpg', and '2020 Analysis.jpg' contain sample outputs when this script is 
	run on the provided data for this exercise.