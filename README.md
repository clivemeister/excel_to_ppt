Scans through an Excel file (default: Insights.xlsx), looking in 
particular columns for keywords and sentiment, as well as for other
data such as type of briefing (partner, customer, etc), location of
briefing, and so on.  Compiles the results into graphics which it
writes out into a Powerpoint file (name is based on the month for
which we are doing the analysis).  

Uses as an input template the file *GCA_Customer_Insights_Month_Year.pptx*

Uses as a source for more sophisticated parameter options the
file *excel_to_ppt.ini*

Key parameters can be put on the command line.  So to run in verbose
mode, for month 8 for 2018, you would use:
    python excel_to_ppt.py -m8 -y2018 -v

This would yield *GCA_Customer_Insights_August-2018.pptx*


