# Linear Regression w Excel (toolpak/manual)

Uses the 23/11/2021 .csv dataset from OWID: COVID 19 Data:
[https://github.com/owid/covid-19-data]

## Current Features
- Select different coefficients using Training and Test sets
- Predict future deaths using Analysis Toolpak 


## Instructions for MacOS/no ToolPak
• Copy CellRange Main!D11:D14 and paste their values in Formulas!D27:D30<br />
• Access Data Analysis Toolpak through ([Data] tab > [Data Analysis] > [Regression])<br />
• Fill in the following:<br />
&nbsp;&nbsp;&nbsp;&nbsp;– Input Y Range: Formulas!$O$2:$O$6652<br />
&nbsp;&nbsp;&nbsp;&nbsp;– Input X Range: Formulas!$K$2:$N$6652<br />
&nbsp;&nbsp;&nbsp;&nbsp;– Check "Labels"<br />
&nbsp;&nbsp;&nbsp;&nbsp;– Output Range: Formulas!$T$3<br />
    
_then_

• Go to Sheet “Formulas (SplitData)” and on CellRange C7, type “training”.<br />
• Now go to Sheet “DataSetSort” and on Column S (“Rand No”), sort the values in the column from lowest to highest.<br />
• This can be done by clicking the “Filter” icon next to the “Rand No” Header and selecting “Sort Smallest to Largest”.<br />
• Now remove the “training” text from CellRange C7 from Sheet “Formulas (SplitData)”.<br />
• Access Data Analysis Toolpak through ( [Data] tab > [Data Analysis] > [Regression] )<br />
• Fill in the following:<br />
&nbsp;&nbsp;&nbsp;&nbsp;– Input Y Range: ‘Formulas (SplitData)’!$O$2 :$O$4657<br />
&nbsp;&nbsp;&nbsp;&nbsp;– Input X Range: ‘Formulas (SplitData)’!$K$2 :$N$4657<br />
&nbsp;&nbsp;&nbsp;&nbsp;– Check “Labels”<br />
&nbsp;&nbsp;&nbsp;&nbsp;– Output Range; ‘Formulas (SplitData)’!$T$3<br />
    
    

## To-Do
- Set up Manual Linear Regression Calculator.
- Import Git CSV and transform automatically.
