# Medical-Data-Analysis
JSON Metadata Extraction and Statistical Analysis

This repository contains a Python script to read JSON files, extract metadata, perform statistical analysis, and create a PowerPoint presentation with the results. The script is designed to analyze medical image metadata to calculate the Insall-Salvati Index (IS ratio), perform outlier detection, and conduct correlation and t-test analyses.

Prerequisites

Make sure you have the following packages installed:

bash
Copy code
pip install json pandas matplotlib seaborn scipy python-pptx
File Description

main.py
This is the main script that performs the following tasks:

Read JSON Files: Extract metadata from JSON files.
Analyze Data: Calculate descriptive statistics, detect outliers, and visualize the data.
Merge DataFrames: Combine the extracted metadata with patient information.
Perform T-Test: Conduct a two-sample t-test to compare IS ratios between male and female patients.
Explore Relationships: Analyze the correlation between IS ratio and other variables.
Create Presentation: Generate a PowerPoint presentation summarizing the analysis results.
Functions
read_json_files(files): Reads JSON files and extracts metadata.
analyze_data(df): Performs data analysis, including descriptive statistics and outlier detection.
merge_dataframes(df1, df2): Merges two dataframes based on specified columns.
perform_two_sample_ttest(merged_df, alpha=0.05): Performs a two-sample t-test on IS ratios.
explore_relationships(merged_df): Analyzes the correlation between IS ratio and other variables.
create_presentation(desc_stats, outliers, correlation_results, t_test_results): Creates a PowerPoint presentation with the analysis results.
Usage

Define JSON Files: Update the jfiles variable in the main function with the paths to your JSON files.
Run the Script: Execute the script using Python.
bash
Copy code
python main.py
Output Files:
Task_A_cleaned_combined_measurements_and_IS_ratio.xlsx: Cleaned data with IS ratios.
Task_B_descriptive_statistics_after_outliers_removal.csv: Descriptive statistics after outlier removal.
merged_data_with_patient_info.xlsx: Merged data with patient information.
correlation_matrix_merged_data.xlsx: Correlation matrix of the merged data.
filtered_correlation_matrix_merged_data.xlsx: Filtered correlation matrix.
boxplot_is_ratio.png: Boxplot of IS ratios.
histogram_is_ratio.png: Histogram of IS ratios.
filtered_correlation_matrix_heatmap.png: Heatmap of the filtered correlation matrix.
Task_C_statistical_analysis_results.pptx: PowerPoint presentation summarizing the analysis results.
Example

The main function is the entry point of the script. It reads the JSON files, performs the analysis, and creates the presentation. Here's an example of how to use the script:

python
Copy code
if __name__ == "__main__":
    main()
This script is designed to automate the process of extracting and analyzing metadata from JSON files, providing a comprehensive statistical analysis and visualization of the results.
