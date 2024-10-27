# Test Scores Analysis Excel Project

This project involves calculating and displaying overall scores for candidates, recording macros for sorting and analysis, and assigning these macros to buttons for easy access. The final workbook provides a streamlined way to organize and analyze candidate data based on scores, location, and birth year.

## Features

1. **Calculate the Overall Score**:
   - Each candidate’s overall score is calculated based on their test score.
   - Additional points are awarded as follows:
     - **+10 points** if the candidate is an employee of Austin city.
     - **+5 points** if the candidate was born in the 1970s.
   - The calculated overall score will appear in the designated column.

2. **Record Macros in the Workbook**:
   - **Sort_score Macro**: Sorts the candidates in the "Candidates" worksheet by the overall score in descending order.
   - **Sort_name Macro**: Sorts the candidates first by Last Name and then by First Name in ascending order.
   - **Pivot Table and Chart Macro**: Inserts a pivot table to calculate the average score for each state and generates a chart to visualize this data.

3. **Assign Macros to Buttons or Shapes**:
   - Adds interactive elements to the worksheet:
     - **Sort_score Button/Shape**: Triggers the `Sort_score` macro.
     - **Sort_name Button/Shape**: Triggers the `Sort_name` macro.
     - **Pivot Table and Chart Button/Shape**: Triggers the Pivot Table and Chart macro for quick data analysis.

4. **Save the Workbook**:
   - Save the workbook as a **macro-enabled file** (.xlsm) to ensure all macros function properly.

## How to Use

1. **Open the Workbook**:
   - Open the `scores.xlsm` workbook in Excel to view and interact with the data.

2. **Calculate Overall Scores**:
   - Review the calculated overall scores in the appropriate column for each candidate.

3. **Use Macro Buttons**:
   - Click on the "Sort by Score" button to sort candidates by overall score in descending order.
   - Click on the "Sort by Name" button to organize candidates alphabetically by Last and First Name.
   - Use the "Pivot & Chart" button to generate a pivot table and chart that displays the average score per state.

4. **Save Changes**:
   - Remember to save any updates or new entries to the workbook in `.xlsm` format to retain macro functionality.


