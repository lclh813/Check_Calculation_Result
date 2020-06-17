# Check Calculation Result

## Spyder Project
<br>
<div align=center><img src="https://github.com/lclh813/Check_Calculation_Result/blob/master/0_Pic/P_0_Project_Structure.png"/></div>
<br>

- ```data_config.py``` Define constants.
- ```query_db.py``` Extract original data stored in SQL.
- ```query_excel.py``` Extract calculated results from Excel.
- ```check_excel_db.py``` Make comparison between SQL and Excel.
- ```check_excels.py``` Make comparison between Excel worksheets.
- ```line_msg.py``` Send Line messages when checking is completed.
- ```main.py``` Contain all the execution codes.

## Jupyter Notebook
It is possible that GitHub fails to display Jupyter Notebooks. Should such circumstances arise, please refer to ***Part 4. Steps*** listed below for code samples.

## Part 1. Objective
To validate the report by:
- Ensuring the original data is correctly imported.
- Making compariosn between calculated results run by VBA with those run by Excel formula.
- Verifying the calculated results are located at the desired Excel cells.

## Part 2. Data
### 2.1. Data 1
- Create database based on which the calculations will be done.
- Tool: ```SQLite```  

### 2.2. Data 2
- Retrieve calculated results from Excel to make further comparison. 
- Tool: ```xlwings```

## Part 3. Outline
### 3.1. Check Original Data
- To verify if original data is imported into Excel correctly by comparing data retreived from Database and Excel.
- Tool: ```SQLite```  ```xlwings```

### 3.2. Check Calculation
- To check the validity of VBA script by comparing with the calculated results run by Excel formula.
- Tool: ```xlwings```

### 3.3. Check Data Location 
- To check numbers of given category are shown at the desired cells.
- Tool: ```xlwings```

## Part 4. Steps
> [***Complete Code***](https://nbviewer.jupyter.org/github/lclh813/Check_Calculation_Result/blob/master/5_CompleteCode.ipynb) 

#### [Step 1. Preparation](https://nbviewer.jupyter.org/github/lclh813/Check_Calculation_Result/blob/master/1_Preparation.ipynb) 

#### [Step 2. Check Original Data](https://nbviewer.jupyter.org/github/lclh813/Check_Calculation_Result/blob/master/2_CheckOriginalData.ipynb) 

#### [Step 3. Check Calculation](https://nbviewer.jupyter.org/github/lclh813/Check_Calculation_Result/blob/master/3_CheckCalculation.ipynb)

#### [Step 4. Check Data Location](https://nbviewer.jupyter.org/github/lclh813/Check_Calculation_Result/blob/master/4_CheckDataLocation.ipynb)
