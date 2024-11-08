# NYC Airbnb Data Cleaning Project

This project involves cleaning and preparing an Airbnb dataset from various neighborhoods in New York City (Queens, Manhattan, Staten Island, Brooklyn, and the Bronx) to ensure data accuracy and consistency for analysis. The project was completed using Microsoft Excel functions and VBA scripting to automate and standardize key cleaning steps.

## Project Overview
The goal of this project was to clean and prepare an Airbnb dataset, focusing on maintaining data integrity, handling missing values, removing duplicates, standardizing the format, and detecting outliers. By improving data quality, this project lays the foundation for reliable analysis and insights into Airbnb listings across NYC.

## Data
  **Dataset**: NYC Airbnb dataset for selected neighborhoods.
  **Source**: Public data sourced from Kaggle.
  **Key Columns**:iD,Name,Host ID, Host Name,Neighborhood Group,Neighborhood,Latitude,Longitude,Room Type,Price,Minimum Nights,Number of Review, Reviews Per Month,Calculated Host Listing Count Availability_365.

## Steps and Methods

### 1. Data Integrity
   **Objective**: Ensured the dataset was accurate, consistent, and reliable.
     **Approach**: Used Excel functions and VBA to check for inconsistencies and validate data entries. Each field was cross-checked to ensure it met the expected standards for analysis.

### 2. Missing Data Handling
   **Objective**: Addressed gaps in the data by imputing or handling missing values as needed.
    **Approach**: Employed Excel functions like `IF`, `ISBLANK`, and conditional formatting to identify and treat missing values. VBA was used to automate imputation where applicable.

### 3. Duplicate Removal
   **Objective**: Removed duplicate records to maintain data uniqueness and avoid bias.
     **Approach**: Applied Excel’s `Remove Duplicates` feature and VBA scripting to detect and eliminate duplicates, ensuring each entry represented a unique listing.

### 4. Standardization
  **Objective**: Ensured consistent formatting and units across the dataset to enable accurate analysis.
     **Approach**: Standardized text, date, and currency formats using Excel functions like `TEXT` and `VALUE`, and automated repetitive tasks with VBA for efficiency.

### 5. Outlier Detection
   **Objective**: Identified and addressed outliers to prevent skewed analysis or model performance.
    **Approach**: Calculated upper and lower bound and used conditional formatting to highlight values beyound and above lower bound and upper bound respectively. Outliers were assessed for relevance, and 'if' function was used to keep to make a decision on how to handle the outliers.
## Tools & Technologies
  **Microsoft Excel**: Used for data functions, formulas, and data validation
  **VBA (Visual Basic for Applications)**: Automated data cleaning tasks for efficiency

## Key Findings
 Identified and managed missing data and outliers across multiple NYC neighborhoods.
  Standardized key columns to ensure consistent units and formats.
  Achieved a clean and reliable dataset, improving the accuracy and reliability for subsequent analysis.

## Instructions for Use
To replicate this project:
1. Download the dataset from [Kaggle](https://www.kaggle.com) or any other source.
2. Open the dataset in Excel and ensure `VBA` is enabled.
3. Run the VBA script provided in this repository to automate cleaning tasks.

## Project Insights
This project highlighted the importance of data quality in analysis. By using Excel and VBA to clean the dataset, potential biases and inaccuracies were minimized, setting up a solid foundation for analysis into Airbnb trends across NYC.

---

**Next Steps**: Continue with additional analyses such as pricing trends by neighborhood, average availability, and customer reviews. For further insights, consider adding visualizations to capture the dataset’s cleaned structure.
