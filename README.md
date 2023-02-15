# Reddit-Scrapper-Finalized

All interaction with this files should be doing within lines 238 to 263

    # ======================================================================
    # Get submission by the provided thread ID.
    threadIDs = ["tmwyx5", "zkveqe"]

    # Will stop scraping once the post being processed is older than this timestamp.
    dt = datetime(2021, 1, 17, 5, 46)

    # Sheet Page name to append to
    page_name = "22.12.1 12-13-12xx"

    # True = Update all the up and downvotes for each of the comments listed
    update_votes = False
 
    # True = Update all the comments/replies that are past the stated data time
    update_commets = True
    # Note: This will not preserve anything that is not past the stated "dt" 

    # True = Append a all values that are not in the excel sheet that are within the "dt" time frame to the bottom of the sheet
    append_comments = False

    # True = Generate and save wordcloud
    generate_wordcloud = True

    # True = Generate and save sentiment graph 
    generate_sentimentGraph = True
    # ======================================================================