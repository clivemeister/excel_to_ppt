import pandas as pd
from datetime import datetime, date
import collections
from collections import Counter
import calendar
import os
import sys
import logging
import re
import sys, getopt

logger = logging.getLogger(__name__)
console=logging.StreamHandler()
console.setLevel(logging.WARNING)
formatter=logging.Formatter('%(asctime)s %(levelname)s %(message)s')
console.setFormatter(formatter)
logger.addHandler(console)

yyyy,mm=date.today().year,date.today().month

def print_help():
    print("count_word.py --w=<word>            count occurences of this word in Insights.xlsx")
    print("              --year=<yyyy>         year to process, default is current year")
    print("              --month=<mm>          month to process, default is current month")
    print("              -d                    turn on debugging trace")
    print("              -i                    turn on information trace")
    return

word_to_find="synergy"
excel_file="Insights.xlsx"

if __name__=="__main__":
    logging.debug("Parsing arguments")
    try:
        opts, args = getopt.getopt(sys.argv[1:],"hid",["w=","year=","month="])
    except getopt.GetoptError:
        print_help()
        sys.exit(2)
    for opt, arg in opts:
        if opt == "-h":
            print_help()
            sys.exit()
        elif opt == "-d":
            ##console.setLevel(logging.DEBUG)
            logger.setLevel(logging.DEBUG)
        elif opt == "-i":
            ##console.setLevel(logging.INFO)
            logger.setLevel(logging.INFO)
        elif opt in ("--w"):
            word_to_find = arg
        elif opt in ("--year"):
            logging.debug("Found argument yyyy with {}".format(arg))
            yyyy = int(arg)
        elif opt in ("--month"):
            logging.debug("Found argument mm with {}".format(arg))
            mm = int(arg)

def make_date(cell_val):
    """Used to create a better-structured date value from the sometimes-odd values in the date column
       Returns a datetime
    """
    if type(cell_val) is datetime:
        v=cell_val.date()
    elif type(cell_val) is str:
        try:
            v=datetime.strptime(cell_val,"%b %d, %Y").date()
        except ValueError:
            v=date.fromordinal(1)
    else:
        v=date.fromordinal(2)
    return v

def tidy_text(cell_val):
    """Standardises the text in a cell:
       Returns the tidied text
    """
    if type(cell_val) is str:
        cell = cell_val.lower()
        cell=re.sub('[\n&@,.:-]',' ',cell)
        cell=" ".join(cell.split())   # idiom to turn multiple spaces between words into single spaces
        for k,v in synonym_list.items():
            cell=cell.replace(k,v)
    else:
        cell=""
    return cell

def bool_list_of_occurrences(series,kwd):
    """Return a boolean list, 1 where an element of 'series' contains word 'kwd', else 0
       Returns a list of 1s and 0s
    """
    pattern = r'\b{0}\b'.format(kwd)
    bool_list=series.str.contains(pattern)
    return bool_list

def count_rows_with_comments(df):
    """ Count the number of rows in dataframe with a comment in either <Want to Learn More About> or <Action Items>
        Returns the count
    """
    return (df["Want to Learn More About"].notnull() | df["Action Items"].notnull()).sum()

def dataframe_for_month(df, year=2018, month=1):
    """Yields a subset the dataframe with only those rows in the given month
       Returns a dataframe
    """
    mm=month
    yyyy=year
    if (mm==12): mm2,yyyy2=1,yyyy+1
    else: mm2,yyyy2=mm+1,yyyy

    logger.debug("Selecting rows for %i-%i" % (yyyy,mm))
    month_df = df.loc[ (df.date>=date(yyyy,mm,1)) & (df.date<date(yyyy2,mm2,1)) ]
    logger.debug("Found %i rows for this month" % len(month_df.index))

    return month_df

def replace_strings(series,repl_dict):
    for k,v in repl_dict.items():
        series=series.str.replace(k,v)
    return(series)

def count_word_usage(df, kwd):
    """Passed a dataframe with columns of at least 'wtlma' and 'ai', together with a keyword
       Return a list (actually a Series) with True wherever the dataframe contains kwd or a synonym in the relevant columns,
       and False otherwise.  Can be used to subset the original dataframe to pick out the rows with the keyword in them:
           df[found_word_list(df,"foo")]   => subset of df with "foo" in one of the columns
    """
    strings_wtlma = replace_strings(df.wtlma, synonym_list)
    strings_ai = replace_strings(df.ai, synonym_list)
    found_list = bool_list_of_occurrences(strings_wtlma,kwd) | bool_list_of_occurrences(strings_ai,kwd)
    return sum(found_list)

logger.info('Starting run for %i-%i...' % (yyyy,mm))

# Read in the ini file
import configparser
import json
ini_file = 'excel_to_ppt.ini'
cfg = configparser.ConfigParser()
cfg.optionxform = str    # read strings as-is from INI file (default is to lowercase keys)
cfg.read(ini_file)

# load the list of synonyms (or mis-spellings) of the keywords
synonym_list={}
for i in cfg.items('synonyms'):
    lst=json.loads(i[1])
    for j in lst:
        synonym_list[j]=i[0]


logger.info("Importing excel file "+excel_file)
all_df = pd.read_excel(open(excel_file,'rb'),header=8,usecols="A:S")

logger.debug("Adding structured columns")
all_df.insert(loc=0,column='date',value=all_df['Visit Date'].apply(make_date))
all_df.insert(loc=1,column='wtlma',value=all_df['Want to Learn More About'].apply(tidy_text))
all_df.insert(loc=2,column='ai',value=all_df['Action Items'].apply(tidy_text))

## Starting 6 months back, count how often the keyword appears in each month

# calc month after this one, to set upper limit for search
if (mm==12): mm_end,yyyy_end=1,yyyy+1
else: mm_end,yyyy_end=mm+1,yyyy
# calc 6 months ago, to set lower limit for search
if (mm<6): mm_start,yyyy_start=mm+7,yyyy-1
else: mm_start,yyyy_start=mm-5,yyyy

word_percent=[0,0,0,0,0,0]
for i in range(0,6):
    this_yyyy=yyyy_start
    this_mm=mm_start+i
    if (this_mm>12):
        this_mm=this_mm-12
        this_yyyy=this_yyyy+1
    df_month = dataframe_for_month(all_df, year=this_yyyy, month=this_mm)
    rows_with_comments=count_rows_with_comments(df_month)
    rows_with_word=count_word_usage(df_month,word_to_find)
    word_percent[i]=rows_with_word/rows_with_comments

print("For last 6 months <%s> usage is: " %(word_to_find))
print(word_percent)
