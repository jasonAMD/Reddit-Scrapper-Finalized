# Reddit-Scrapper-Finalized

## Set up
Run the following command in the terminal window:
    pip install -r requirements.txt

If this gives an error then try:
    pip install -r "{}""
Where {} is equal to the directory that you have placed the folder in so for example this may look like:
    pip install -r "C:\Users\jasokhuu\Desktop\GitHub Code\RedditScraper\2022 Vanguard Reddit Defect Tracker"

## All interaction with this files should be doing within lines 238 to 291

    # ======================================================================
    # Get submission by the provided thread ID.
    threadIDs = ["zkveqe", "vod0y7"] Insert any additional threads into this list 

    # Will stop scraping once the post being processed is older than this timestamp.
    dt = datetime(2022, 1, 17, 5, 46)

    # Insert the name of the excel file
    name_excelFile = "2022 Vanguard Reddit Defect Tracker.xlsx"

    # Sheet Page name to work on
    page_name = "22.12.1 12-13-12xx"

    # True = Update all the up and downvotes for each of the comments listed
    update_votes = False
 
    # True = Update all the comments/replies that are past the stated data time
    update_commets = False
    # Note: This will not preserve anything that is not past the stated "dt" 

    # True = Append a all values that are not in the excel sheet that are within the "dt" time frame to the bottom of the sheet
    append_comments = True

    # True = Generate and save wordcloud
    generate_wordcloud = False

    # True = Generate and save sentiment graph 
    generate_sentimentGraph = False
    # ======================================================================