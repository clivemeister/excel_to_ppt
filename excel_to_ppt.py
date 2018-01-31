import pandas as pd
from datetime import datetime, date
import matplotlib.pyplot as plt
import collections
from collections import Counter
import calendar
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Pt
import os
import sys
import logging
import re

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

console=logging.StreamHandler()
console.setLevel(logging.INFO)
formatter=logging.Formatter('%(asctime)s %(levelname)s %(message)s')
console.setFormatter(formatter)
logger.addHandler(console)

yyyy,mm=2017,12
logger.info('Starting run for %i-%i...' % (yyyy,mm))

# Read in the ini file
import configparser
import json
ini_file = 'excel_to_ppt.ini'
cfg = configparser.ConfigParser()
cfg.read(ini_file)

colour_list=[]   # list of the colour code values
if 'colours' in cfg:
    colour_codes = dict(cfg.items('colours'))
    for i in cfg.items('colours'):
        colour_list.append(i[1])       # put this hex value in list of available colours
else:
    logger.error('No [colours] section in {}'.format(ini_file))

if 'keywords' in cfg:
    c_to_k=cfg.items('keywords')
    dict_colour_of_keywords={}
    for i in cfg.items('keywords'):
        if i[0] in colour_codes:
            dict_colour_of_keywords[colour_codes[i[0]]]=json.loads(i[1])
        else:
            logger.error('In {}, section [keywords], found colour {} which was not listed in [colours] section'.format(ini_file,i[0]))
else:
    logger.error('No [keywords] section in {}'.format(ini_file))

#use all the things we have a colour for as the vocab to look for
vocab=[]
for i in dict_colour_of_keywords.values():
    for j in i:
        vocab.append(j)

# load the list of synonyms (or mis-spellings) of the keywords
synonym_list={}
for i in cfg.items('synonyms'):
    lst=json.loads(i[1])
    for j in lst:
        synonym_list[j]=i[0]

industry_list=["China","CME","Energy","Fin Svcs","Health & LS",
            "Japan", "Mfg","Other","RCG","Public Sector", "Travel & Trans"]
centres = ["PA","LON1","SNG","NY1","H"]

def make_date(cell_val):
    """Used to create a better-structured date value from the sometimes-odd values in the date column
       Returns a datetime
    """
    if type(cell_val) is datetime:
        v=cell_val.date()
    elif type(cell_val) is str:
        try:
            v= datetime.strptime(cell_val,"%b %d, %Y").date()
        except ValueError:
            v=date.fromordinal(1)
    else:
        v=date.fromordinal(2)
    return v

def replace_strings(series,repl_dict):
    for k,v in repl_dict.items():
        series.str.replace(k,v)
    return(series)

def keyword_counts_in_series(series,keywords):
    list_of_strings = replace_strings(series.str.lower().str.replace('[\n&,.-]',' '),
                             synonym_list
                            )
    if keywords is str:
        #get here only if keywords is a single string, not a list
        count_list=sum(list_of_strings.str.find(keywords)>=0)
    else:
        count_list=Counter()
        for i in keywords:
            count_list[i]=sum(list_of_strings.str.find(i)>=0)
    return count_list


def dataframe_for_month(df, year=2017, month=1):
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

def dataframe_for_6months(df, year=2017, month=1):
    """Yields a subset the dataframe with those rows for last 6 months up to given month
       Returns a dataframe
    """
    mm=month
    yyyy=year
    # calc month after this one, to set upper limit for search
    if (mm==12): mm_end,yyyy_end=1,yyyy+1
    else: mm_end,yyyy_end=mm+1,yyyy
    # calc 6 months ago, to set lower limit for search
    if (mm<=6): mm_start,yyyy_start=mm+6,yyyy-1
    else: mm_start,yyyy_start=mm-6,yyyy

    logger.debug("Selecting rows for %i-%i to %i-%i" % (yyyy_start,mm_start,yyyy,mm))
    month_df = df.loc[ (df.date>=date(yyyy_start,mm_start,1)) & (df.date<date(yyyy_end,mm_end,1)) ]
    logger.debug("Found %i rows for this month" % len(month_df.index))

    return month_df

def keywords_in_dataframe(df):
    """ Count each keyword in the dataframe's important columns
        Returns a Counter collection of (keyword:count)
    """
    import operator   #for sorted(), used in sorting items into OrderedDict

    logger.debug("Counting occurrences of keywords for this month")
    c_wtlma = keyword_counts_in_series(df['Want to Learn More About'],vocab)
    c_actions=keyword_counts_in_series(df['Action Items'],vocab)

    logger.debug("Adding together the results")
    counter_words_to_freq = c_wtlma + c_actions

    return counter_words_to_freq

def count_rows_with_comments(df):
    """ Count the number of rows in dataframe with a comment in either <Want to Learn More About> or <Action Items>
        Returns the count
    """
    return (df["Want to Learn More About"].notnull() | df["Action Items"].notnull()).sum()

def counts_by_centre(dataframe):
    centre_counts = dataframe.Ctr.value_counts()
    counts_list = (centre_counts.PA if hasattr(centre_counts,"PA") else 0,
                   centre_counts.H if hasattr(centre_counts,"H") else 0,
                   centre_counts.NY1 if hasattr(centre_counts,"NY1") else 0,
                   centre_counts.LON1 if hasattr(centre_counts,"LON1") else 0,
                   centre_counts.SNG if hasattr(centre_counts,"SNG") else 0
                   )
    return counts_list

class SimpleGroupedColorFunc(object):
    """Create a color function object which assigns EXACT colors
       to certain words based on the color to words mapping

       Parameters
       ----------
       color_to_words : dict(str -> list(str))
         A dictionary that maps a color to the list of words.

       default_color : str
         Color that will be assigned to a word that's not a member
         of any value from color_to_words.
    """

    def __init__(self, color_to_words, default_color):
        self.word_to_color = {word: color
                              for (color, words) in color_to_words.items()
                              for word in words}

        self.default_color = default_color

    def __call__(self, word, **kwargs):
        return self.word_to_color.get(word, self.default_color)

class GroupedColorFunc(object):
    """Create a color function object which assigns DIFFERENT SHADES of
       specified colors to certain words based on the color to words mapping.

       Uses wordcloud.get_single_color_func

       Parameters
       ----------
       color_to_words : dict(str -> list(str))
         A dictionary that maps a color to the list of words.

       default_color : str
         Color that will be assigned to a word that's not a member
         of any value from color_to_words.
    """

    def __init__(self, color_to_words, default_color):
        from wordcloud import get_single_color_func
        self.color_func_to_words = [
            (get_single_color_func(color), set(words))
            for (color, words) in color_to_words.items()]

        self.default_color_func = get_single_color_func(default_color)

    def get_color_func(self, word):
        """Returns a single_color_func associated with the word"""
        try:
            color_func = next(
                color_func for (color_func, words) in self.color_func_to_words
                if word in words)
        except StopIteration:
            color_func = self.default_color_func

        return color_func

    def __call__(self, word, **kwargs):
        return self.get_color_func(word)(word, **kwargs)

def file_wordcloud_for_month(keywords_for_month, useful_rows_for_month, year=2017, month=8):
    """Input is a keywords_for_month, a Counter, plus useful_rows_for_month, an integer, and year and month as integers
       Returns the filename where the wordcloud is stored
    """
    ##Build a wordcloud, using the wordcloud code from Andreas Mueller
    # (to install, run "pip install wordcloud"
    from wordcloud import WordCloud
    import matplotlib.pyplot as plt
    import os

    # set font so 30% occurrence of top word uses 196 point font
    percent = round(100*(keywords_for_month.most_common(1)[0][1]/useful_rows_for_month))
    font_for_biggest_word = round( 196 * percent/30 )
    logger.info("font for word {} is {} based on {}% = {} / {}".format(
                               keywords_for_month.most_common(1)[0][0],
                                     font_for_biggest_word,
                                             percent,
                                                keywords_for_month.most_common(1)[0][1],
                                                        useful_rows_for_month))

    font_path = os.path.join("C:",os.sep,"Windows",os.sep,"Fonts",os.sep,'arial.ttf')
    logger.debug("Generating the wordcloud")
    wc_for_month = WordCloud(font_path=font_path,
                             width=2500,height=500,
                             prefer_horizontal=1.0,
                             relative_scaling=0.7,
                             max_font_size=font_for_biggest_word,
                             background_color="white",
                             random_state=1
                             ).generate_from_frequencies(keywords_for_month)


    # Words that are not in any of the dict_colour_of_keywords values
    # will be colored with a grey single color function
    default_color = 'grey'

    # Apply our color function
    #grouped_color_func = GroupedColorFunc(dict_colour_of_keywords, default_color)
    grouped_color_func = SimpleGroupedColorFunc(dict_colour_of_keywords, default_color)
    wc_for_month.recolor(color_func=grouped_color_func)

    # Build the generated image
    plt.imshow(wc_for_month,interpolation='bilinear')
    plt.axis("off")
    plt.figure()
    #plt.show()

    filename = "wordcloud-"+calendar.month_name[month]+".png"
    plt.imsave(filename,wc_for_month,format="png")
    plt.close()

    return filename


def file_graph_for_month_kwd(kwd,kwd_pos,vals,months,line_color):
    fig,ax=plt.subplots(figsize=(5.75,3.25))
    ##############ax.clear()

    ax.plot(vals,line_color,linewidth=3)
    ax.set_axis_off()
    plt.ylim(min(vals)-0.05,max(vals)+0.05)
    m_minus_2_percent = "{0:.0f}%".format(vals[0] * 100)
    m_this_percent    = "{0:.0f}%".format(vals[2] * 100)
    ax.text(0,vals[0],calendar.month_name[months[0]]+"\n"+m_minus_2_percent,fontsize=36)
    ax.text(2,vals[2],calendar.month_name[months[2]]+"\n"+m_this_percent,fontsize=36)
    filename = "graph-"+str(kwd_pos)+".png"
    logger.debug("Saving %s graph for keyword %s in file %s" % (kwd_pos,kwd,filename))
    fig.savefig(filename,bbox_inches="tight")
    plt.close(fig)

    return filename

def file_donut_pie_for_month(values,kwd):
    fig, ax = plt.subplots()
    ax.axis('equal')
    outside, _ = ax.pie(values,startangle=90,counterclock=False,
                        colors=list(colour_list))
    ax.legend(("Palo Alto","Houston","NY","London","Singapore"),fontsize=24,bbox_to_anchor=(0.8,1.0),frameon=False)
    width = 0.50  #determines the thickness of donut rim
    plt.setp( outside, width=width, edgecolor='white')

    filename = "donut-"+re.sub(r"[& ]","_",str(kwd))+".png"
    logger.debug("Saving %s donut in file %s with values %r" % (kwd, filename, values))
    fig.savefig(filename,bbox_inches="tight")

    plt.close(fig)

    return filename

def file_donut_pie_for_industries(industries):
    fig, ax = plt.subplots()
    ax.axis('equal')
    width = 2.0
    outside, _ = ax.pie(industries.values,
                        startangle=90,counterclock=False,
                        colors=list(colour_list),
                        radius=5.0)
    ax.legend(industries.index.tolist(),ncol=3,fontsize=24,loc=10,bbox_to_anchor=(0.5,-2),frameon=False)
    plt.setp( outside, width=width, edgecolor='white')

    filename = "donut-industries.png"
    logger.debug("Saving Industries donut in file %s with values %r" % (filename, industries))
    fig.savefig(filename,bbox_inches="tight")

    plt.close(fig)

    return filename

def write_customer_list(df_for_kwd,text_frame):
    text_frame.text = "Customers"
    font=text_frame.paragraphs[0].font
    font.size = Pt(12)
    font.bold = True
    font.color.theme_color = MSO_THEME_COLOR.DARK_1

    for i in (df_for_kwd)["Account Name"].str.title().iteritems():
        new_para = text_frame.add_paragraph()
        font=new_para.font
        font.size = Pt(8)
        font.color.theme_color = MSO_THEME_COLOR.DARK_2
        new_para.text=i[1]

    return

def write_top_keywords(text_frame,titleText,kwd_counts_for_ind,rowcount):
    """Write a section headed by titleText, with bullets of the top 4 elements of Counter kwd_counts_for_ind plus
       their percentage of use, where the percentage is calculated with rowcount as the denominator.
         e.g:
           c=Counter([("alpha":5),("bravo":4),("charlie":3),("delta":2))])
           write_top_keywords(tf,"My list",c,10)
         yields:
           My list
            - alpha - 50%
            - bravo - 40%
            - charlie - 30%
    """
    text_frame.text = titleText
    font=text_frame.paragraphs[0].font
    font.size = Pt(10)
    font.bold = True
    font.color.theme_color = MSO_THEME_COLOR.DARK_1

    for k,v in kwd_counts_for_ind.most_common(4):
        new_para = text_frame.add_paragraph()
        font=new_para.font
        font.size = Pt(8)
        font.color.theme_color = MSO_THEME_COLOR.DARK_2
        new_para.text=" - "+k+" - "+"{0:.0f}%".format(v/rowcount * 100)

    return

def find_text_in_shapes(slide_shapes,searchword):
    """Find the given text in the list of powerpoint shapes passed in slide_shapes
    """
    found_idx = -1
    for i in range(len(slide_shapes)):
        if (slide_shapes[i].has_text_frame):
            text_frame = slide_shapes[i].text_frame
            if (text_frame.text.find(searchword)>=0):
                found_idx = i
                break

    return found_idx

def replace_text_in_shape(slide_shapes,find,use,slidename):
    i = find_text_in_shapes(slide_shapes,find)
    if (i>=0):
        text_frame = slide_shapes[i].text_frame
        text_frame.text = use
    else:
        logger.error("Could not find %s placeholder on %s" % (find,slidename))

    return

def new_run_in_slide(para,text='text',fontname='Arial',fontsize=24):
    "Add new run to a paragraph in a powerpoint slide"
    run = para.add_run()
    run.text = text
    font = run.font
    font.size = Pt(fontsize)
    font.name = fontname
    return run

excel_file='Insights.xlsx'
logger.info("Importing excel file "+excel_file)
all_df = pd.read_excel(open(excel_file,'rb'),header=8)

logger.debug("Adding structured date column")
all_df.insert(loc=0,column='date',value=all_df['Visit Date'].apply(make_date))

## Given the month and year, calc the number of the previous few months
if (mm==1): mm_minus_1,year_for_mm_minus_1 = 12, yyyy-1
else: mm_minus_1, year_for_mm_minus_1 = mm-1, yyyy
if (mm<=2): mm_minus_2,year_for_mm_minus_2 = mm+10, yyyy-1
else: mm_minus_2, year_for_mm_minus_2 = mm-2, yyyy
if (mm<=6): mm_minus_6,year_for_mm_minus_6 = mm+6, yyyy-1
else: mm_minus_6, year_for_mm_minus_6 = mm-6, yyyy

### Now start to generate the powerpoint
from pptx import Presentation
from pptx.util import Inches, Pt

# These appear to be the layouts used in the master for this slide deck
LAYOUT_TITLE_WITH_DARK_PICTURE = 0
LAYOUT_TITLE_SLIDE_WITH_NAME   = 1
LAYOUT_DIVIDER                 = 2
LAYOUT_TITLE_ONLY              = 3
LAYOUT_TITLE_AND_SUBTITLE      = 4
LAYOUT_BLANK                   = 5

## Open up the source presentation
prs = Presentation('GCA_Customer_Insights_Month-Year.pptx')
this_month = calendar.month_name[mm]
earliest_month = calendar.month_name[mm_minus_2]
logger.info("Creating presentation for %s" % (this_month))

################################################
## Modify existing title slide with month
################################################
logger.info("Modifying text in slide 0")
s = prs.slides[0]
slide_shapes = s.shapes
text_frame = slide_shapes[0].text_frame  # should be the the title textframe
#clear existing text, and write new text into title textframe
text_frame.clear()
# First para with a run of 60-point Arial font
new_run_in_slide(text_frame.paragraphs[0],text='Customer Insights',fontname='Arial',fontsize=60)
# Second para with a run of 32-point font
new_run_in_slide(text_frame.add_paragraph(),text='Learnings from '+this_month+' EBC/CEC visits',fontsize=32)

################################################
## 3rd slide: count the keywords for this month and build the wordcloud
################################################
logger.info("3rd slide: large wordcloud for this month, and top 3 keywords")
df_for_month = dataframe_for_month(all_df, year=yyyy, month=mm)
kwd_count_for_month = keywords_in_dataframe(df_for_month)
useful_rows_in_m = count_rows_with_comments(df_for_month)
logger.info("Top keyword/counts for month %i : %r" % (mm,kwd_count_for_month.most_common(5)) )
file_wordcloud_for_month(kwd_count_for_month, useful_rows_in_m,
                         year=yyyy,month=mm)
## Modifying main wordcloud slide by changing title and adding pic for this month's wordcloud
logger.info("Modifying text and adding wordcloud in slide 2")
s = prs.slides[2]
slide_shapes=s.shapes
title_frame = slide_shapes.title.text_frame
title_frame.clear()
new_run_in_slide(title_frame.paragraphs[0],text="In "+this_month+" customers wanted to learn more about...",
       fontname="Arial",fontsize=28)
left=Inches(0.5); top=Inches(4.0)
slide_shapes.add_picture("wordcloud-"+this_month+".png",left,top,height=Inches(2.5))

##Find where the placeholders are for the top 3 keywords update them
top_3 = kwd_count_for_month.most_common(3)   # top 3 keywords for most recent month in list with their counts
kwd0=top_3[0][0]
kwd1=top_3[1][0]
kwd2=top_3[2][0]
logger.debug("Top 3 keywords for current month are %s %s %s" % (kwd0,kwd1,kwd2))
replace_text_in_shape(slide_shapes,find="topword_1",use=kwd0,slidename="3rd slide")
replace_text_in_shape(slide_shapes,find="topword_2",use=kwd1,slidename="3rd slide")
replace_text_in_shape(slide_shapes,find="topword_3",use=kwd2,slidename="3rd slide")

################################################
## 4th slide: count the keywords for previous two months, and build their wordclouds.
################################################
logger.info("4th slide: three wordclouds for most recent 3 months")
df_for_month_minus_1 = dataframe_for_month(all_df, year=year_for_mm_minus_1, month=mm_minus_1)
kwd_count_for_m_minus_1 = keywords_in_dataframe(df_for_month_minus_1)
logger.info("Top keyword/counts for month %i : %r" % (mm_minus_1,kwd_count_for_m_minus_1.most_common(5)) )

df_for_month_minus_2 = dataframe_for_month(all_df, year=year_for_mm_minus_2, month=mm_minus_2)
kwd_count_for_m_minus_2 = keywords_in_dataframe(df_for_month_minus_2)
logger.info("Top keyword/counts for month %i : %r" % (mm_minus_2,kwd_count_for_m_minus_2.most_common(5)) )

useful_rows_in_m_2 = count_rows_with_comments(df_for_month_minus_2)
useful_rows_in_m_1 = count_rows_with_comments(df_for_month_minus_1)
logger.info("Useful rows in months 2,1,0 are %i, %i, %i" % (useful_rows_in_m_2,useful_rows_in_m_1,useful_rows_in_m))

file_wordcloud_for_month(kwd_count_for_m_minus_1, useful_rows_in_m_1, year=year_for_mm_minus_1,month=mm_minus_1)
file_wordcloud_for_month(kwd_count_for_m_minus_2, useful_rows_in_m_2, year=year_for_mm_minus_2,month=mm_minus_2)

## Modifying 3-month wordcloud slide by adding pic for last 3 months' wordcloud
logger.info("Adding three wordclouds in slide 3")
s = prs.slides[3]
slide_shapes=s.shapes
left=Inches(3.65)
top=Inches(1.4)
slide_shapes.add_picture("wordcloud-"+calendar.month_name[mm_minus_2]+".png",left,top,height=Inches(1.5))
top=Inches(3.25)
slide_shapes.add_picture("wordcloud-"+calendar.month_name[mm_minus_1]+".png",
                         left,top,height=Inches(1.5))
top=Inches(5.05)
slide_shapes.add_picture("wordcloud-"+this_month+".png",
                         left,top,height=Inches(1.5))
#Find where the placeholders are for the keywords whose frequency we are graphing and update them
replace_text_in_shape(slide_shapes,find="Month-2",use=calendar.month_name[mm_minus_2],slidename="4th slide")
replace_text_in_shape(slide_shapes,find="Month-1",use=calendar.month_name[mm_minus_1],slidename="4th slide")
replace_text_in_shape(slide_shapes,find="Month-0",use=calendar.month_name[mm],slidename="4th slide")

################################################
## 5th slide: for top 3 keywords for this month graph their usage
################################################
logger.info("5th slide: top keywords for this month")
months=[mm_minus_2,mm_minus_1,mm]

kwd0_c2 = kwd_count_for_m_minus_2[kwd0] if (kwd0 in kwd_count_for_m_minus_2) else 0
kwd0_c1 = kwd_count_for_m_minus_1[kwd0] if (kwd0 in kwd_count_for_m_minus_1) else 0
kwd0_c0 = kwd_count_for_month[kwd0]     # must have keyword as it came from this dictionary
vals_kwd0=[kwd0_c2/useful_rows_in_m_2, kwd0_c1/useful_rows_in_m_1, kwd0_c0/useful_rows_in_m]
file_linegraph_topic1 = file_graph_for_month_kwd(kwd0,"1st",vals_kwd0,months,colour_list[0])
logger.info("Kwd0 is %s, data %r" % (kwd0,vals_kwd0))
kwd1_c2 = kwd_count_for_m_minus_2[kwd1] if (kwd1 in kwd_count_for_m_minus_2) else 0

kwd1_c1 = kwd_count_for_m_minus_1[kwd1] if (kwd1 in kwd_count_for_m_minus_1) else 0
kwd1_c0 = kwd_count_for_month[kwd1]     # must have keyword as it came from this dictionary
vals_kwd1=[kwd1_c2/useful_rows_in_m_2, kwd1_c1/useful_rows_in_m_1, kwd1_c0/useful_rows_in_m]
file_linegraph_topic2 = file_graph_for_month_kwd(kwd1,"2nd",vals_kwd1,months,colour_list[1])
logger.info("Kwd1 is %s, data %r" % (kwd1,vals_kwd1))

kwd2_c2 = kwd_count_for_m_minus_2[kwd2] if (kwd2 in kwd_count_for_m_minus_2) else 0
kwd2_c1 = kwd_count_for_m_minus_1[kwd2] if (kwd2 in kwd_count_for_m_minus_1) else 0
kwd2_c0 = kwd_count_for_month[kwd2]     # must have keyword as it came from this dictionary
vals_kwd2=[kwd2_c2/useful_rows_in_m_2, kwd2_c1/useful_rows_in_m_1, kwd2_c0/useful_rows_in_m]
file_linegraph_topic3 = file_graph_for_month_kwd(kwd2,"3rd",vals_kwd2,months,colour_list[2])
logger.info("Kwd0 is %s, data %r" % (kwd2,vals_kwd2))

## Build a subset of the dataframe for last 3 months that uses each of the top 3 kwds in this month
df_for_3months = pd.concat([df_for_month,df_for_month_minus_1,df_for_month_minus_2])
df_for_kwd0 = df_for_3months.loc[(df_for_3months['Want to Learn More About'].str.lower().str.find(kwd0)>=0) |
                                 (df_for_3months['Action Items'].str.lower().str.find(kwd0)>=0)
                                ]
df_for_kwd1 = df_for_3months.loc[(df_for_3months['Want to Learn More About'].str.lower().str.find(kwd1)>=0) |
                                 (df_for_3months['Action Items'].str.lower().str.find(kwd1)>=0)
                                ]
df_for_kwd2 = df_for_3months.loc[(df_for_3months['Want to Learn More About'].str.lower().str.find(kwd2)>=0) |
                                 (df_for_3months['Action Items'].str.lower().str.find(kwd2)>=0)
                                ]

## Create donut pies showing split of visits expressing interest in top 3 topics by centre over last 3 months
kwd0_counts = counts_by_centre(df_for_kwd0)
file_donut_topic1 = file_donut_pie_for_month(counts_by_centre(df_for_kwd0),"1st")
file_donut_topic2 = file_donut_pie_for_month(counts_by_centre(df_for_kwd1),"2nd")
file_donut_topic3 = file_donut_pie_for_month(counts_by_centre(df_for_kwd2),"3rd")

## Add the line graphs and donut pies to the Top 3 Customer Interests chart
logger.info("Adding line graphs, donuts and customers to the Top 3 Customer Interests slide (5th slide)")
s = prs.slides[4]
slide_shapes=s.shapes
#Update the title
title_frame = slide_shapes.title.text_frame
title_frame.clear()
new_run_in_slide(title_frame.paragraphs[0],text="Top 3 Customer Interests: "+earliest_month+"-"+this_month,
       fontname="Arial",fontsize=28)
#Find where the placeholders are for the keywords whose frequency we are graphing and update them
replace_text_in_shape(slide_shapes,find="Topic1",use=kwd0,slidename="5th slide")
replace_text_in_shape(slide_shapes,find="Topic2",use=kwd1,slidename="5th slide")
replace_text_in_shape(slide_shapes,find="Topic3",use=kwd2,slidename="5th slide")

#Add the line graphs and the donuts for each of the topics
top=Inches(1.8); h=Inches(1.0); w=Inches(1.7)
slide_shapes.add_picture(file_linegraph_topic1,Inches(0.7),top,height=h,width=w)
slide_shapes.add_picture(file_linegraph_topic2,Inches(4.7),top,height=h,width=w)
slide_shapes.add_picture(file_linegraph_topic3,Inches(8.7),top,height=h,width=w)
top=Inches(1.7); h=Inches(1.1); w=Inches(2.25);
slide_shapes.add_picture(file_donut_topic1,Inches(2.3),top,height=h,width=w)
slide_shapes.add_picture(file_donut_topic2,Inches(6.35),top,height=h,width=w)
slide_shapes.add_picture(file_donut_topic3,Inches(10.35),top,height=h,width=w)

#Find where the placeholders are for the customer lists and update them
df_for_kwd0_month0 = df_for_month.loc[(df_for_month['Want to Learn More About'].str.lower().str.find(kwd0)>=0) |
                                 (df_for_month['Action Items'].str.lower().str.find(kwd0)>=0)
                                ]
df_for_kwd1_month0 = df_for_month.loc[(df_for_month['Want to Learn More About'].str.lower().str.find(kwd1)>=0) |
                                 (df_for_month['Action Items'].str.lower().str.find(kwd1)>=0)
                                ]
df_for_kwd2_month0 = df_for_month.loc[(df_for_month['Want to Learn More About'].str.lower().str.find(kwd2)>=0) |
                                 (df_for_month['Action Items'].str.lower().str.find(kwd2)>=0)
                                ]
i = find_text_in_shapes(slide_shapes,"Customers1")
if (i>=0):
    logger.debug("Writing list of %i customers for first keyword" % (len(df_for_kwd0)))
    write_customer_list(df_for_kwd0_month0,slide_shapes[i].text_frame)
else:
    logger.error("Could not find Customers1 placeholder")

i = find_text_in_shapes(slide_shapes,"Customers2")
if (i>=0):
    logger.debug("Writing list of",len(df_for_kwd1),"customers for second keyword")
    write_customer_list(df_for_kwd1_month0,slide_shapes[i].text_frame)
else:
    logger.error("Could not find Customers2 placeholder")

i = find_text_in_shapes(slide_shapes,"Customers3")
if (i>=0):
    logger.debug("Writing list of",len(df_for_kwd2),"customers for third keyword")
    write_customer_list(df_for_kwd2_month0,slide_shapes[i].text_frame)
else:
    logger.error("Could not find Customers3 placeholder")

################################################
## 9th slide: Now generate the Industry Insights donuts and top keyword lists
################################################
logger.info("9th slide: industries and how they broke down across the centres")
## Generate the image for the big donut showing volume for all industries across the months
industry_counts = df_for_3months["Updated Industry Sep6"].value_counts()
file_donut_ind_vols = file_donut_pie_for_industries(industry_counts)

## Generate the individual donuts showing, for each industry, the spread across the centres


# set up some dictionaries to hold the industry datasets for later
df_for_ind={};
file_donut_for_ind={};
kwd_counts_for_ind={}

for ind in industry_list:
    df_for_ind[ind] = df_for_3months[df_for_3months["Updated Industry Sep6"]==ind]
    file_donut_for_ind[ind] = file_donut_pie_for_month(counts_by_centre(df_for_ind[ind]),ind)
    kwd_counts_for_ind[ind] = keywords_in_dataframe(df_for_ind[ind])
    logger.info("Top keyword/counts for %s: %r" % (ind,kwd_counts_for_ind[ind].most_common(4)) )

## Now write out the charts and text on industry (9th) slide
s = prs.slides[8]
slide_shapes=s.shapes

#Update the title
title_frame = slide_shapes.title.text_frame
title_frame.clear()
new_run_in_slide(title_frame.paragraphs[0],text="Industry Insights "+earliest_month+"-"+this_month,
       fontname="Arial",fontsize=28)

#Add the one big donut showing volumes for each industry to the slide
logger.info("Adding main donut to Industry Insights slide (9th slide)")
left=Inches(1.7); top=Inches(1.9)
slide_shapes.add_picture(file_donut_ind_vols,left,top,height=Inches(4.5),width=Inches(3.7))

#Add the individual industry donuts broken down by centre to the slide
#Need to pick the top 5, excluding "Other", from the list in industry_list
logger.info("Adding donuts for top industries in each centre (9th slide)")
left=Inches(8.2); top=(Inches(1.20),Inches(2.30),Inches(3.40),Inches(4.50),Inches(5.60))
h=Inches(1.0); w=Inches(2.05)
n=0 # used to count the number of industry pies we have placed (we can't use enumerate(industry_counts) as we don't always place a pie)
for ind in industry_counts.index:
    if   (ind!="Other"):
        logger.info("Writing %s as industry %i" %(ind,n))
        #Write the industry name as the main title for this box
        replace_text_in_shape(slide_shapes,"Industry-{}".format(n),ind,"9th slide")
        #Add the donut showing breakdown of centres that hosted this industry
        slide_shapes.add_picture(file_donut_for_ind[ind], left,top[n], height=h,width=w)
        #Now write the list of top interests for this industry
        idx = find_text_in_shapes(slide_shapes,"Top interests - {}".format(n))
        if (idx>=0):
            text_frame = slide_shapes[idx].text_frame
            write_top_keywords(text_frame,"Top Interests - "+ind,
                               kwd_counts_for_ind[ind],
                               count_rows_with_comments(df_for_ind[ind])
                              )
        else:
            logger.error("Could not find <Top interests - {}> placeholder on 9th slide".format(n))
        n+=1
    if (n>=5): break    # Stop after we've put 5 pictures in place

################################################
## 10th slide: Partner insights
################################################
logger.info("10th slide: top partner keywords and partner attendance broken out by centre")
# rather copmlex expression to find which rows are partners who are attending as partners, or partner-led briefings for customers
df_for_partners = df_for_3months.loc[ ( df_for_3months['Account Type'].isin(["Channel/ Reseller","Systems Integrator"])
                                        & (df_for_3months['Partner / Customer']!="Customer")
                                      ) | (df_for_3months['Partner / Customer']=="Partner")
                                    ]
kwd_count_for_partners = keywords_in_dataframe(df_for_partners)
useful_rows_in_partners = count_rows_with_comments(df_for_partners)

df_channel = df_for_partners[ df_for_partners['Account Type']=="Channel/ Reseller" ]
df_SI = df_for_partners[ df_for_partners['Account Type']=="Systems Integrator" ]
df_accompanied = df_for_partners[ df_for_partners['Partner / Customer']=="Partner" ]


SI_count=[len(df_SI[df_SI.Ctr==c]) for c in centres]
Channel_count=[len(df_channel[df_channel.Ctr==c]) for c in centres]
Attended_count=[len(df_accompanied[df_accompanied.Ctr==c]) for c in centres]
logger.info("Centres being looked at: %r" %(centres))
logger.info("Volume by centre - partner: %r" %(Channel_count))
logger.info("Volume by centre - SI: %r" %(SI_count))
logger.info("Volume by centre - attended: %r" %(Attended_count))

fig, ax = plt.subplots()
fig.set_size_inches(5.5,2.2)
y=(0.5,1,1.5,2,2.5)
ax.barh(y,Channel_count, height=0.35, color=colour_list[0],tick_label=centres)
ax.barh(y,SI_count, height=0.35, left=Channel_count,color=colour_list[1])
new_base=[x+y for x,y in zip(Channel_count, SI_count)]
ax.barh(y,Attended_count, height=0.35, left=new_base,color=colour_list[2])
ax.legend(["Channel/ Reseller","SI","Partner attended with customer"],loc='lower center',ncol=3,bbox_to_anchor=(0.5,-0.4))
ax.get_xaxis().set_visible(False)
file_hbar = "barh-PartnerVolumes.png"
fig.savefig(file_hbar,bbox_inches="tight")

plt.close(fig)

## Start to update the slide
s = prs.slides[9]
slide_shapes=s.shapes

#Update the title
title_frame = slide_shapes.title.text_frame
title_frame.clear()
new_run_in_slide(title_frame.paragraphs[0],text="Partner Insights ("+earliest_month+"-"+this_month+")",
       fontname="Arial",fontsize=28)

# Update the top 4 most common keywords, and their percentages
logger.info("top 4 partner interests: %r" % (kwd_count_for_partners.most_common(4)))
for n,p in enumerate(kwd_count_for_partners.most_common(4)):
    # p is (keyword: count) for each of the top 4 most common keywords
    replace_text_in_shape(slide_shapes,"Interest-{}".format(n),p[0],"10th slide")
    replace_text_in_shape(slide_shapes,"Score-{}".format(n),"{0:.0f}%".format(100*p[1]/useful_rows_in_partners),"10th slide")

# Place the horizontal bar graph
slide_shapes.add_picture(file_hbar, Inches(0.9),Inches(4.4), height=Inches(2.2),width=Inches(5.5))

################################################
## 11th slide: Top interests and industries for last 6 months
################################################
logger.info("11th slide: top 5 interests, top 3 industries, and their top interests, by centre, for last 6 months")
df_6months = dataframe_for_6months(all_df, year=yyyy, month=mm)

#build a dictionary of dataframes subsetted by centre, a dictionary of keywords by centre,
#and a dict of top 3 industries per centre and their keywords (where the key is a tuple of (centre, industry) )
dfs_6m_ctr={}
kwd_counts_6m_ctr={}
industry_counts_6m = df_6months["Updated Industry Sep6"].value_counts()
kwd_counts_6m_for_top_inds={}
commented_rows_for_6m={}

for c in centres:
    dfs_6m_ctr[c]=df_6months[df_6months.Ctr==c]
    kwd_counts_6m_ctr[c]=keywords_in_dataframe(dfs_6m_ctr[c])
    industry_counts_6m[c] = (dfs_6m_ctr[c])["Updated Industry Sep6"].value_counts()
    for ind in (industry_counts_6m[c]).index:
        this_df = dfs_6m_ctr[c][ dfs_6m_ctr[c]["Updated Industry Sep6"]==ind ]
        kwd_counts_6m_for_top_inds[(c,ind)]=keywords_in_dataframe( this_df )
        commented_rows_for_6m[(c,ind)]=count_rows_with_comments(this_df)
        logger.info("Top keyword/counts in %s for %s: %r" % (c,ind,kwd_counts_6m_for_top_inds[(c,ind)].most_common(3)) )


## Build the slide: add the interests, industries, and per-industry interests, for each of the centres
logger.info("Setting the title ")
s = prs.slides[10]
slide_shapes=s.shapes
#Update the title
title_frame = slide_shapes.title.text_frame
title_frame.clear()
new_run_in_slide(title_frame.paragraphs[0],text="Breakdown by centre ("+calendar.month_name[mm_minus_6]+"-"+this_month+")",
       fontname="Arial",fontsize=28)

#Update, by centre, the top interests, and the top industries with their top interests
logger.info("Setting the per-centre top 5 interests, together with per-centre top 3 industries and their interests")
for c in centres:

    #First, the top interests for this centre
    for n,p in enumerate(kwd_counts_6m_ctr[c].most_common(5)):
        # p is (keyword: count) for each of the keywords, so p[0] is the keyword itself
        replace_text_in_shape(slide_shapes,"{}-interest-{}".format(c,n),p[0],"11th slide")

    #Next, the top industries with their interests for this centre - industry_counts_6m[c] is already ordered highest->lowest count
    n=0  #count how many displayed - need to do this separately from the loop count, as we ignore "Other" as an industry group
    for ind in industry_counts_6m[c].index:
        if   (ind!="Other"):
            logger.info("For centre <%s> industry <%i> is <%s>" %(c,n,ind))
            #Write the list of top interests for this industry
            idx = find_text_in_shapes(slide_shapes,"{}-industry-{}".format(c,n))
            if (idx>=0):
                write_top_keywords(slide_shapes[idx].text_frame,
                                   ind,
                                   kwd_counts_6m_for_top_inds[(c,ind)],
                                   commented_rows_for_6m[(c,ind)]
                                  )
            else:
                logger.error("Could not find <{}-industry-{}> placeholder on 11th slide".format(c,n))
            n+=1   #increment the number of industries written out for this centre
        if (n>=3): break   #exit the loop after writing in 3 industries+interests


################################################
## Close the source presentations
################################################
logger.info("Saving Powerpoint file for "+this_month)
prs.save('GCA_Customer_Insights_'+this_month+'-'+str(yyyy)+'.pptx')
## Close any open figures
plt.close("all")
logger.info("...and we're done!")
for h in list(logger.handlers): logger.removeHandler(h)   # may be several here if we've crashed sometimes
