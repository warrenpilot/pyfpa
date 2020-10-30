# pyfpa

This project provides the basic Financial, Planning & Analysis functions in Python.  The functions include:

- Collecting data from Excel files into Pandas Dataframes
- Ability to map custom budget or actual reports to capture data and dimensions
- Creating a data cube repository for financial, operational, sales or any kind of data.  A Golden Source for actuals, budgets, forecasts, sales reports, web statistics, stock data ...anything you can think of!
- Source and version control to keep track of which files are the basis for data
- Easy slicing and dicing of data based on dimensions you define
- Changing table data into record data for pivot tables
- Dimension management to accomodate changes
- Consolidation based on dimesions
- Variance analysis
- Pasting back into Excel
- Providing a basis to use all of Python's data science tools

The goal is to make an easier introduction into Python for finance analysts who do the work of collecting and analyzing data.  Python, and especially Pandas, can be daunting for uses in FP&A, but does provide advantages:

- Can handle large data and associated calculations
- No incorrect links or changing underlying data
- API connections to almost every type of database and software
- Access to a greater amount of data science tools (statistical, AI)
- Access to high-end charting and visualization tools
- And all for *free*

While overpowered for most FP&A functions (Excel is a great tool), this package looks to leverage that power to address the challenges of FP&A activities.