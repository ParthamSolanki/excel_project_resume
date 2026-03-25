# Excel project with basic dashboard and visualizations

In this project I will be cleaning a dataset, adding some useful columns, replacing rows with one value to other values, then creating KPI sheets, pivot tables and visualizations using this data. Using these visualizations I will create a dashboard to track all the stuff from a centralized place. Since the data is of employees, it will end up looking like an HR analytics project with a basic dashboard.

## Basic cleaning and data transformation with some useful columns added.

- **Creating copies of the raw data** - Creating a copy of the original file, or saving as an excel workbook to create multiple sheets within one as csv doesn't support that then creating a copy of the sheet in the workbook and using that without touching the original as say if something goes wrong, I can roll back. Doing both, creating a copy of the data as well as creating a copy of the sheet in the excel workbook itself.

- **Changing formatting of the textual stuff** - If you want to change the formatting of the names and textual stuff and feel like it is not properly formatted (this dataset feels fine when I glossed over it), you can just use UPPER, LOWER or PROPER to formate the stuff in a standardized format. You can also remove the additional columns by copying the values over the original column to make the whole sheet look neater.

- **Making the sheet look neat** - Most of the columns look congested and some of the data is not showing properly so I will quickly select all the data and use the `Autofit column width` function in the Format section in the Home pane.

- **Removing duplicates** - Selecting all the columns and removing duplicates using the in-built function of Excel.
    No duplicates found.

- **Glossing over the columns to remove anomalies** - Adding up filters for viewing to see their are any anomalies in the data or some stuff needs changing, then using this will view most of the columns as needed.
    - It seems that there is no managerID for the manager named "Webster Butler". After checking the manager name column using the filter dropdown, it seems that some of the other columns have the managerID for this manager which is 39. So we will populate all the other empty cells for this person with this managerID.

    Seems like no other column has any problems with NULL values of empty cells.
    
- **Changing blanks into NULL** - There are many in the Date of termination but that's to be expected as it would be empty for people who haven't been terminated. 
    For changing these blank cells to NULL -> Select the column > Go to special from the upper ribbon > Toggle Blank on then press ok > Type NULL without clicking anywhere > Ctrl + Enter to populate all the blank cells.
 Since its advisable to change it to NULL or something like that so that it doesn't cause issues later on when you are working with the data in a querying language, so we will be changing it. It can also be done manually by populating all the empty cells that show after you filter for blank cells.

- **TRIM unnecessary spaces** - Since while looking through the columns there was not any problem with the leading or trailing spaces, so will not be trimming but if needed can be done using the trim function promptly and easily. While making the KPI sheet I noticed that there are some trailing space after Production department, so going to tackle that and replace the original column.

- **Formatting and assigning correct datatype to dates** - Standardizing the dates to be of the date datatype or if there is any sort of problem or inconsistency with the way they are written it will be corrected using the inbuilt datatype changing function of excel after selecting all the columns. Will still recheck using the filter to see if it messed something up and take care of it.

- **Separating the whole name into first and last name columns** - Will be creating some useful columns like first name and last name as the names are in one column.
    - For separating the first name to its own column use `=TRIM(MID(A2, FIND(",", A2)+1, LEN(A2)))` this will use the MID to extract all the stuff after the ',' and then trim it as there is a space after the ',' but if you are sure that there is a space in all the rows you can just use `=MID(A2, FIND(",", A2)+2, LEN(A2))` instead. Mine contains some rows where there is no space after the delimiter so I will just use the previous one.
    - For separating the last name to its own column use `=LEFT(A2, FIND(",", A2)-1)`, and then auto populate all the other rows in the columns.

- **Creating an age column to get the age of the employees in years** - The DOB has been formatted properly so we can just use `=DATEDIF(R2, TODAY(), "Y")` to get the present age of individuals dynamically. We can also check the minimum and maximum ages using the filter which are 33 and 74 respectively.

- **Creating an age bracket column to make more insightful visualizations** - We are going to use the column of age created in the previous step to segregate the people into age brackets of say 10 year intervals. The code will be `=IF(AND(AM2>=30, AM2<=40), "30-40", IF(AND(AM2>40, AM2<=50), "40-50", IF(AND(AM2>50, AM2<=60), "50-60", IF(AND(AM2>60, AM2<=70), "60-70", IF(AM2>70, "70+ years", "NULL")))))` . It can seem pretty hectic but its just a basic nested IF as long as you remember the brackets and don't input too many arguments into one IF statement.

- **Creating an Employee Attrition column** - There is already an Employment status column, but creating an attrition column with only yes and no values can simplify it and make it more usable for dashboards and calculating attrition rates. Created using this function in an empty column `=IF(Y2="NULL", "No", "Yes")`

- **Finding the names and ids of managers** - I tried to match if the manager were also listed as employees, which was not the case, if they were I would have created an Ismanager column where yes or no would have told if that person is a manager or not. Alternatively I am just going to copy the manager name and ids to a different sheet and use remove duplicates to get the list of managers.

- **Creating an additional column for count of employees each** - We will count all the employees that are managed by each manager to get a sense of who is managing a ton or people and who is managing less. Use this function in an empty cell and then auto populate all the below rows - `=COUNTIF(Working!AD:AD, Managers!A2)`

- **Changing the genderID column to gender** - The genders of the employees are listed as 1 for male and 2 for female, will be changing it to make it more presentable by selecting the column and finding and replacing.

## Creating a KPI sheet

- Creating a heading for the sections by combining and centering multiple cells
    - Workforce Overview

- **Creating an active employee cell** - Use `=COUNTIF(Working!M:M, 0)`

- **Creating a pivot table for attrition rate by department**
    - Rows - Department_Fixed
    - Values - Employee count, Attrition count, Average of Employee Satisfaction, Average of absences, Average of salary, Attrition Rate
While making the pivot table using the department, I noticed that there are trailing spaces in the Production row value, so we are going to remove them and replace the original row.
Adding a slicer based on gender will make it more useful as well as there won't be a need to make another pivot table based on gender.

- **Creating a metrics table based on age brackets sliced over gender** - This will give an overview of the stuff based on the age brackets of the people which can be further focused based on genders.

Also have an additional sheet for pivot tables with more parameters named Pivot tables. It includes pivot tables for
- Attrition rate by department
- Attrition rate by gender
- Metrics per department based on gender
- Metrics based on age brackets sliced over gender

## Creating a basic interactive dashboard

- Creating a copy of the pivot tables sheet to simplify the pivot tables and create more simplistic tables for making visualizations for the dashboard.

- Most of the visualizations will be used from the data derived in the KPI sheet.

- First we will make the heading after combining some cells and add active and terminated employees alongside it for quick viewing.

- Making a gender column for only active individuals in a new sheet to make a gender split.

- Adding visualizations of
    - Employee count by department
    - Average salaries by department with slicer for gender
    - Attrition rate by department
    - Metrics like Count, Attrition, Average number of special projects based on age brackets.

Ending the project for creating the dashboard here as it feels pretty usable and I have already included most of the common stuff that is required, other stuff is ever so slightly more niche and without any context there is no need to keep working on it.