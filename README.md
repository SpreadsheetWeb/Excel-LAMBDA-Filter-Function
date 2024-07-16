# Excel Lambda Function for Search Operations

## Overview

This repository contains an Excel LAMBDA function designed to perform multi-criteria search operations on a given data table. This function, named `FilterLambda`, allows users to search across three different columns and filter data based on the provided search terms from specific cells.

## Function Definition

```excel
=LAMBDA(SearchText1; SearchText2; SearchText3; DataTable; ErrorText; NoDataText;
  LET(
    FilteredData1; IF(ISBLANK(SearchText1); DataTable;
      FILTER(DataTable; ISNUMBER(SEARCH(UPPER(SearchText1); UPPER(INDEX(DataTable;;1))))));
    FilteredData2; IF(ISBLANK(SearchText2); FilteredData1;
      FILTER(FilteredData1; ISNUMBER(SEARCH(UPPER(SearchText2); UPPER(INDEX(FilteredData1;;2))))));
    FilteredData3; IF(ISBLANK(SearchText3); FilteredData2;
      FILTER(FilteredData2; ISNUMBER(SEARCH(UPPER(SearchText3); UPPER(INDEX(FilteredData2;;3))))));
    Result; IFERROR(FilteredData3; ErrorText);
    IF(ROWS(Result) = 0; NoDataText; Result)
  )
)
```

### Parameters

- **SearchText1**: The first search term to filter the data based on the first column.
- **SearchText2**: The second search term to filter the data based on the second column.
- **SearchText3**: The third search term to filter the data based on the third column.
- **DataTable**: The data table to be searched.
- **ErrorText**: The text to be displayed in case of an error.
- **NoDataText**: The text to be displayed if no data matches the search criteria.

## How to Use the Function

1. **Define the LAMBDA Function in Excel:**
   - Go to the `Formulas` tab.
   - Click on `Name Manager` and then click on `New`.
   - In the `Name` field, enter `FilterLambda`.
   - In the `Refers to` field, paste the LAMBDA function definition above.
   - Click `OK` to save the function.

2. **Prepare Your Data Table:**
   - Ensure your data is in a table format. For example, name your table `Table1`.
   - Your table should have at least three columns, as the function filters based on the first three columns.

3. **Set Up Search Criteria:**
   - Enter your search criteria in cells `A1`, `A2`, and `A3`. These cells will be used to input the text you want to search for in the respective columns of your table.

4. **Use the LAMBDA Function:**
   - In the cell where you want to display the filtered results, enter the following formula:
   ```excel
   =FilterLambda(A1, A2, A3, Table1, "No Data", "Error")
   ```

### Example

Suppose you have the following data in a table named `Table1`:

| Name    | Department | Location     |
|---------|------------|--------------|
| Alice   | HR         | New York     |
| Bob     | IT         | San Francisco|
| Charlie | Finance    | New York     |
| Dave    | IT         | New York     |
| Eve     | HR         | San Francisco|

To filter this data based on your search criteria, follow these steps:

1. Enter your search criteria:
   - Cell `A1`: "HR"
   - Cell `A2`: "San Francisco"
   - Cell `A3`: Leave blank (if no third search criteria)

2. Use the LAMBDA function at any cell:
   ```excel
   =FilterLambda(A1, A2, A3, Table1, "No Data", "Error")
   ```
The result will be:

| Name    | Department | Location     |
|---------|------------|--------------|
| Eve     | HR         | San Francisco|

### Notes

- The search is case-insensitive.
- If a search text cell is left blank, the function will ignore that search criteria.
- The function will display "No Data" if no matching rows are found.
- The function will display "Error" if an error occurs during filtering.

## Conclusion

This LAMBDA function is a powerful tool to perform dynamic and multi-criteria search operations in Excel. By following the steps outlined above, you can easily integrate and use this function in your own Excel workbooks using specific reference cells for search terms.

## License

This project is licensed under the MIT License.
