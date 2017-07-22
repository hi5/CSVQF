# CSV Quick Filter (CSVQF) - v0.49q

This program allows you to load a CSV file (any delimited file) and use various search criteria to
filter the listview.  
You can export the results to a new file.  
The regular expression search is case sensitive and should be a Perl-compatible regular expression
(PCRE, www.pcre.org)

**Note:** an entire row of the CSV is searched at once and not on a cell by cell basis to provide
faster search results.

## Command line options

    CSVQF file ["delimiter"] ["header"] ["Columns to use in CSV"]

Example, opening a | delimited file:

    CSVQF data.csv "|"

Use the first row as header for listview: (header = 1,Y,Yes,T,True)

    CSVQF data.csv "," "1"

Use \t for a tab delimited file

    CSVQF tabdata.txt "\t"

Use \s for a space delimited file

    CSVQF data.txt "\s"

Only use specific columns (1 & 5):

    CSVQF data.csv "," "0" "1,5"

0 here means don't use first row as header

## Searching

There are several search options available in the GUI. To search in a specific column you can use
a shorthand like so:

    2|there

means search for 'there' in the second column, you can combine several column searches:

    2|there 5|that

means search for 'there' in the second column and 'that' in the fitfh column.

If the number of columns in irregular, you can select "LAST" in the column pulldown menu and it will
always search the last column that has data in it.

## Feedback

Visit the forum at https://autohotkey.com/boards/viewtopic.php?f=6&t=34867    
You can post as guest, no registration required.

This program is written in AutoHotkey, a free, open-source (script) utility for Windows, you can learn more at www.autohotkey.com

## Screenshot

![CSV Quick Filter - CSVQF - window](https://raw.github.com/hi5/CSVQF/master/img/csvqf-inaction.gif)

Scroll through the results with the (Page)Up/(Page)Down keys, view row data with Enter.

### Credits

* CSV Library [lib] by trueski and kdoske - https://github.com/hi5/CSV
* OnChangeMyText by jsherk - https://autohotkey.com/board/topic/68189-validate-characters-as-you-type-in-an-edit-control/
* Attach by majkinetor - https://autohotkey.com/board/topic/49214-ahk-ahk-l-forms-framework-08/

Archived forum thread:

* https://autohotkey.com/board/topic/68279-csv-quick-filter-gui-show-results-in-listview-as-you-type/
