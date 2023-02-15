import praw
from datetime import datetime, timezone
import xlwings as xw
import pandas as pd
import re
import string
from textblob import TextBlob
import matplotlib.pyplot as plt
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
from transformers import AutoTokenizer, AutoModelForSequenceClassification
import torch
import matplotlib.pyplot as plt
from wordcloud import WordCloud, STOPWORDS
from xlwings.constants import DeleteShiftDirection

class RedditComment:
    def __init__(self, thread, body, link, user, timeStamp, upVotes, downVotes, commentDepth):
        self.thread = thread
        self.body = body
        self.link = link
        self.user = user
        self.timeStamp = timeStamp
        self.upVotes = upVotes
        self.downVotes = downVotes
        self.commentDepth = commentDepth
        self.bert = "-1"
        self.textblob = "-1"
        self.sentiment = "-1"
        self.accuracy = "-1"

    def __str__(self):
        message = "Thread: {} \nBody: {} \nLink: {} \nuser: {} \nTime Stamp: {} \nTime Updated: {} \nUp Votes: {}" \
                  "\nDown Votes: {} \nComment Depth: {} \nBert Score: {} \nTextblob Score: {} \nOverall Sentiment: {}" \
                  "\nAccuracy: {}".format(self.thread, self.body, self.link, self.user, self.timeStamp, self.upVotes,
                                          self.downVotes, self.commentDepth, self.bert, self.textblob, self.sentiment, self.accuracy)
        return message

    def to_dict(self):
        return {'Thread/Comment Body': self.body, 'ASIC': 'IGNORE', 'Applications': 'IGNORE', 'Category': 'IGNORE', 'Ticket ID': 'IGNORE',
                'Notes/Action Items': 'IGNORE', 'URL': self.link, 'ThreadID': self.thread, 'User': self.user, 'UTC Time Posted': self.timeStamp,
                'Upvotes': self.upVotes, 'Downvotes': self.downVotes, 'Comment Depth': self.commentDepth,
                'NLTK/TextBlob Sentiment': self.textblob, 'roBERTa Sentiment': self.bert, 'Average Sentiment': self.sentiment,
                'Accuracy Sentiment Score': self.accuracy
                }

def get_bertScore(comment):
    tokens = tokenizer.encode(comment.body[:512], return_tensors='pt')
    result = model(tokens)
    score = int(torch.argmax(result.logits))
    sentiment = ""
    if score == 0:
        sentiment = 'Negative'
    elif score == 1:
        sentiment = 'Neutral'
    elif score == 2:
        sentiment = 'Positive'
    comment.bert = sentiment


def get_textblobScore(comment):
    sid_obj = SentimentIntensityAnalyzer()
    sentiment_dict = sid_obj.polarity_scores(comment.body)
    scores_dict = sentiment_dict
    if sentiment_dict['compound'] >= 0.2:
        scores_dict['sentiment'] = 'Positive'
    elif sentiment_dict['compound'] <= -0.2:
        scores_dict['sentiment'] = 'Negative'
    else:
        scores_dict['sentiment'] = 'Neutral'
    comment.textblob = scores_dict['sentiment']


def get_finalSentiment(comment):
    full_point, half_point, no_point = 1, 0.5, 0
    bert_score = comment.bert
    blob_score = comment.textblob

    if blob_score == bert_score == "Positive":
        sentiment = "Positive"
        accuracy = full_point
    elif blob_score == bert_score == "Negative":
        sentiment = "Negative"
        accuracy = full_point
    elif blob_score == bert_score == "Neutral":
        sentiment = "Neutral"
        accuracy = full_point
    elif (blob_score == "Positive" and bert_score == "Neutral") or (blob_score == "Neutral" and bert_score == "Positive"):
        sentiment = "Positive"
        accuracy = half_point
    elif (blob_score == "Negative" and bert_score == "Neutral") or (blob_score == "Neutral" and bert_score == "Negative"):
        sentiment = "Negative"
        accuracy = half_point
    elif (blob_score == "Positive" and bert_score == "Negative") or (blob_score == "Negative" and bert_score == "Positive"):
        sentiment = "Unknown"
        accuracy = no_point

    comment.sentiment = sentiment
    comment.accuracy = accuracy


def get_comment_info(thread_name, comment, depth) -> RedditComment:

    return RedditComment(thread_name,
                         comment.body,
                         r"https://www.reddit.com" + comment.permalink,
                         comment.author,
                         datetime.utcfromtimestamp(comment.created_utc).strftime('%Y-%m-%d %H:%M:%S'),
                         comment.ups,
                         comment.downs,
                         depth)

def preOrderTraversal(root_comment, threadID) -> list:
    list_output = list()

    def dfs(node_comment, depth):
        if not node_comment:
            return
        list_output.append({"CommentItem": node_comment, "Depth": depth, "ThreadID": threadID})
        for reply in node_comment.replies:
            dfs(reply, depth=depth+1)

    dfs(root_comment, 0)

    return list_output

def get_excelsheet_df(wb, page_name):
    sheet = wb.sheets[page_name]
    search_value = "Thread/Comment Body"

    used_range = sheet.used_range

    for cell in used_range:
        if cell.value == search_value:
            cuttoff_row = cell.row - 1
            break

    data_below_cutoff = used_range.offset(cuttoff_row, 0).options(
        pd.DataFrame, expand='table').value

    df_excel = pd.DataFrame(data_below_cutoff).reset_index()
    for column in df_excel.columns:
        df_excel[column] = df_excel[column].astype('str')

    return df_excel

def update_excelsheet(wb, page_name, df, row_number=0):
    sheet = wb.sheets[page_name]

    if row_number == 0:
        search_value = "Thread/Comment Body"

        used_range = sheet.used_range

        for cell in used_range:
            if cell.value == search_value:
                start_row = cell.row + 1
                break
    else:
        start_row = row_number

    for row in range(df.shape[0]):
        for col in range(df.shape[1]):
            if (df.iloc[row, col] == 'IGNORE') or (df.iloc[row, col] == 'None'):
                sheet.range((row+start_row, col+1)).clear_contents()
            else:
                sheet.range((row+start_row, col+1)).value = df.iloc[row, col]

def generate_hyperlink(wb, page_name, cell_string):
    sheet = wb.sheets[page_name]

    start_cell = sheet.range(cell_string)
    end_cell = sheet.range(cell_string).end("down")
    link_range = sheet.range(start_cell, end_cell)

    for cell in link_range:
        sheet.range(cell.row, cell.column).add_hyperlink(cell.value)


def get_thread_comments(reddit, dt, threadID):

    submission = reddit.submission(id=threadID)

    submission.comments.replace_more(limit=None)

    list_comments = list()
    for comment in submission.comments.list():
        # Check for timestamp cutoff.
        if (comment.created_utc < dt.replace(tzinfo=timezone.utc).timestamp()):
            print("Comment is older than {0}, skipping comment.".format(
                dt.strftime(r"%m/%d/%Y, %H:%M")))
        else:
            list_comments.append(preOrderTraversal(comment, threadID))

    list_comments = [item for sublist in list_comments for item in sublist]

    # Removing any duplicates
    seen = []
    filtered_comments = []
    for comment in list_comments:
        if comment["CommentItem"].id not in seen:
            filtered_comments.append(comment)
            seen.append(comment["CommentItem"].id)

    print("Done making the filtered comment list for {}".format(threadID))

    return filtered_comments

def get_cellPosition(wb, page_name, search_value, row_value_only = False):
    sheet = wb.sheets[page_name]

    used_range = sheet.used_range
    
    try:
        if row_value_only == False:
            for cell in used_range:
                if cell.value == search_value:
                    row_number = cell.row + 1
                    column_number = cell.column 
                    column_letter = chr(ord('A') + column_number - 1)
                    return_value = str(column_letter) + str(row_number)
                    break
        else:
            for cell in used_range:
                if cell.value == search_value:
                    row_number = cell.row # This is just what row the cell your looking for is located at
                    return_value = int(row_number)
                    break
    except:
        print("You have entered a value that does not exist in the workbook please enter a valid value to locate")
    
    return return_value

def generate_singleValue(wb, page_name, cell, value):
    sheet = wb.sheets[page_name]
    sheet[cell].value = value

if __name__ == "__main__":
    # ======================================================================
    # Get submission by the provided thread ID.
    threadIDs = ["tmwyx5"] # "zkveqe",  

    # Will stop scraping once the post being processed is older than this timestamp.
    dt = datetime(2022, 1, 17, 5, 46)

    # Sheet Page name to append to
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

    # Connect to Reddit API using registered app code.
    reddit = praw.Reddit(client_id='-ZZ5CjZYA2NxLA',
                        client_secret='vqzwJr9AnhswTiz6-pMRV92Re8w',
                        username='AMD_Erik',
                        password=r'3DIslandsLammers',
                        user_agent='scraper')

    print("Logged in successfully as:", reddit.user.me())

    # Initializing pretrained sentiment analyzer model
    tokenizer = AutoTokenizer.from_pretrained('cardiffnlp/twitter-roberta-base-sentiment-latest')
    model = AutoModelForSequenceClassification.from_pretrained('cardiffnlp/twitter-roberta-base-sentiment-latest')

    print("\n ======================")
    print("\nDone initalizing the pretrained model")

    list_comments = list()
    for threadID in threadIDs:
        print("Now Searching Data for the threadID: {}".format(threadID))
        list_comments.append(get_thread_comments(reddit=reddit, dt=dt, threadID=threadID))
    
    list_comments = [item for sublist in list_comments for item in sublist]

    # Turned that list of praw api comments into custom reddit comment objects
    if (len(list_comments) > 0):
        processed_comments = list()
        for comment in list_comments:
            comment_info = get_comment_info(comment['ThreadID'], comment['CommentItem'], comment['Depth'])
            get_bertScore(comment_info)
            get_textblobScore(comment_info)
            get_finalSentiment(comment_info)
            processed_comments.append(comment_info)
        print("Finished retrieving all of the initial reddit data from the forum site")

        # I know this name is confusing but it's just turning the reddit commment objects into mini datafarmes in the list
        df_form_comments = list()
        for comment in processed_comments:
            df_form_comments.append(comment.to_dict())

        # The final dataframe result
        df_comments = pd.DataFrame(df_form_comments)
        for column in df_comments.columns:
            df_comments[column] = df_comments[column].astype('str')
        print("Have completed processing the information into a dataframe")
    else:
        print("=====" * 5)
        print("ERROR! You have collected 0 comments, please insert a different thread name or change your dt minimumn time")
        exit()

    # -----------------------------------

    wb = xw.Book('C:\\Users\\jasokhuu\\Desktop\\GitHub Code\\RedditScraper\\2022 Vanguard Reddit Defect Tracker.xlsx')
    df_excel = get_excelsheet_df(wb=wb, page_name=page_name)

    if (update_commets == True and update_votes == True) or (update_commets == True and append_comments == True) or (update_votes == True and append_comments == True):
        print("Error you have entered 2 or more 'True's for either update_comments, update_votes and append_commends as 'True'. Please only enter one as 'True' and try again.")

    elif (update_commets == False) and (update_votes == False) and (append_comments == False):
        print("Error you have entered update_comments and update_votes as both 'False'. So nothing happened, please only enter one as 'True' and try again.")

    elif update_commets == True:
        print("Setting to Update Comments was selected")
        if len(df_excel) != 0:
            mapping = dict(zip(zip(df_excel['User'], df_excel['UTC Time Posted']),
                               zip(df_excel['ASIC'], df_excel['Applications'], df_excel['Category'], df_excel['Ticket ID'], df_excel['Notes/Action Items'])))

            df_comments[['ASIC', "Applications", "Category", "Ticket ID", "Notes/Action Items"]] = [
                mapping.get((User, Posted), (asic, Applications, Category, Ticket_ID, Notes)) for User, Posted, asic, Applications, Category, Ticket_ID, Notes
                in zip(df_comments['User'], df_comments['UTC Time Posted'], df_comments['ASIC'], df_comments['Applications'], 
                       df_comments['Category'], df_comments['Ticket ID'], df_comments['Notes/Action Items'])]
            
        update_excelsheet(wb=wb, page_name=page_name, df=df_comments)

        current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cell_datetime = get_cellPosition(wb=wb, page_name=page_name, search_value="Last Comment Scrape")
        generate_singleValue(wb=wb, page_name=page_name, cell=cell_datetime, value=current_datetime)

        # Deleting all excess rows 
        if (len(df_comments) < len(df_excel)):
            sheet = wb.sheets[page_name]
            row_header = get_cellPosition(wb=wb, page_name=page_name, search_value="Thread/Comment Body", row_value_only=True)

            row_starting = row_header + len(df_comments) + 1
            row_ending = row_header + len(df_excel) + 10
            print("Clearing from row {} to {}".format(row_starting, row_ending))

            string_delete = "{}:{}".format(str(row_starting), str(row_ending))
            sheet.range(string_delete).api.Delete(DeleteShiftDirection.xlShiftUp) 

        print("The excel workbook was updated with new comments")

        cell_hyperlink = get_cellPosition(wb=wb, page_name=page_name, search_value="URL")
        generate_hyperlink(wb=wb, page_name=page_name, cell_string=cell_hyperlink)
        print("URLs have been turned into hyperlinks")

    elif update_votes == True:
        print("Setting to Update Votes was selected")
        if len(df_excel != 0):
            mapping = dict(zip(zip(df_comments['User'], df_comments['UTC Time Posted']), zip(df_comments['Upvotes'], df_comments['Downvotes'])))

            df_excel[['Upvotes', "Downvotes"]] = [mapping.get((User, Posted), (Upvotes, Downvotes)) for User, Posted, Upvotes, Downvotes
                                                in zip(df_excel['User'], df_excel['UTC Time Posted'], df_excel['Upvotes'], df_excel['Downvotes'])]

            update_excelsheet(wb=wb, page_name=page_name, df=df_excel)

            current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cell_datetime = get_cellPosition(wb=wb, page_name=page_name, search_value="Last Vote Update")
            generate_singleValue(wb=wb, page_name=page_name, cell=cell_datetime, value=current_datetime)
            print("Updated the excel sheet voting information")

            cell_hyperlink = get_cellPosition(wb=wb, page_name=page_name, search_value="URL")
            generate_hyperlink(wb=wb, page_name=page_name, cell_string=cell_hyperlink)
            print("URLs have been turned into hyperlinks")
        else:
            print("There is nothing in the spreadsheet to update, perhaps update the comments first")

    elif append_comments == True:
        print("Setting to append values was selected")
        if len(df_comments) > 0:

            df_append = pd.concat([df_excel, df_comments])
            df_append = df_append.drop_duplicates(subset=['User', 'UTC Time Posted'])

            current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cell_datetime = get_cellPosition(wb=wb, page_name=page_name, search_value="Last Appending")
            generate_singleValue(wb=wb, page_name=page_name, cell=cell_datetime, value=current_datetime)

            update_excelsheet(wb=wb, page_name=page_name, df=df_append)
            print("Finished appending {} new rows of data".format(str(len(df_append) - len(df_excel))))
            
            cell_hyperlink = get_cellPosition(wb=wb, page_name=page_name, search_value="URL")
            generate_hyperlink(wb=wb, page_name=page_name, cell_string=cell_hyperlink)
            print("URLs have been turned into hyperlinks")
        else:
            print("There was nothing to append")


    if generate_wordcloud == True:
        wordcloud = WordCloud(stopwords=STOPWORDS, background_color="black").generate(' '.join(df_comments["Thread/Comment Body"]))

        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis("off")
        plt.show()

        current_datetime = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        wordcloud.to_file("WordCloud({}).png".format(current_datetime))
    
    if generate_sentimentGraph == True:
        sentiment_scores = {'positive': 0, 'negative': 0, 'neutral': 0}
        for x in df_comments.index:
            sentiment = df_comments['Average Sentiment'][x]
            accuracy = float(df_comments['Accuracy Sentiment Score'][x])
            if sentiment == 'Positive':
                sentiment_scores['positive'] += accuracy
            elif sentiment == 'Neutral':
                sentiment_scores['neutral'] += accuracy
            elif sentiment == 'Negative':
                sentiment_scores['negative'] += accuracy

        names = list(sentiment_scores.keys())
        values = list(sentiment_scores.values())

        plt.bar(range(len(sentiment_scores)), values, tick_label=names)

        current_datetime = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        plt.savefig("SentimentGraph({}).png".format(current_datetime))
        plt.show()

    print("Completed")