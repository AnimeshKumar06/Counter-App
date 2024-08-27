import requests
from datetime import datetime
import time
import random
from bs4 import BeautifulSoup as bs
from googlesearch import search
import pandas as pd
import feedparser

# List of User-Agents
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.1 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36",
  
]

# Function to parse date in multiple formats
def parse_date(date_string):
    for fmt in ('%Y-%m-%d', '%Y-%m-%dT%H:%M:%SZ', '%d %b %Y', '%d %B %Y', '%H:%M'):
        try:
            return datetime.strptime(date_string, fmt).strftime('%Y-%m-%d %H:%M:%S')
        except ValueError:
            continue
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

# Function to fetch articles from web search
def get_articles_from_web(keyword, num_results=200):
    links = []
    query = f"{keyword} news"
    start = 0
    while len(links) < num_results:
        try:
            # Perform search
            search_results = search(query, num_results=num_results, pause=2)
            for url in search_results:
                if len(links) >= num_results:
                    break
                links.append(url)
            start += 10
            time.sleep(random.uniform(1, 3))  # Random delay to avoid being blocked
        except Exception as e:
            print(f"Error fetching articles for keyword '{keyword}': {e}")
            break
    return links

# Function to fetch URL with retries and backoff
def fetch_url(url, retries=3, backoff_factor=0.3):
    for attempt in range(retries):
        headers = {
            "User-Agent": random.choice(USER_AGENTS)
        }
        try:
            response = requests.get(url, headers=headers, timeout=5)
            response.raise_for_status()
            return response
        except (requests.exceptions.RequestException, requests.exceptions.HTTPError) as e:
            print(f"Error fetching URL '{url}', attempt {attempt + 1}/{retries}: {e}")
            if attempt < retries - 1:
                time.sleep(backoff_factor * (2 ** attempt))
    return None

# Function to get article details from a web page
def get_article_details_from_web(url):
    response = fetch_url(url)
    if response:
        try:
            soup = bs(response.text, 'html.parser')
            title = soup.title.string if soup.title else 'No title'

            time_published = None
            for meta in soup.find_all('meta'):
                if 'property' in meta.attrs and meta.attrs['property'] in ['article:published_time', 'og:published_time', 'article:modified_time', 'og:modified_time']:
                    time_published = meta.attrs['content']
                    break
                if 'name' in meta.attrs and meta.attrs['name'] in ['pubdate', 'publishdate', 'timestamp', 'DC.date.issued']:
                    time_published = meta.attrs['content']
                    break

            if not time_published:
                time_element = soup.find('time')
                if time_element and 'datetime' in time_element.attrs:
                    time_published = time_element['datetime']

            if not time_published:
                time_published = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            else:
                time_published = datetime.strptime(time_published[:19], '%Y-%m-%dT%H:%M:%S').strftime('%Y-%m-%d %H:%M:%S')

            paragraphs = soup.find_all('p')
            content = ' '.join([para.get_text() for para in paragraphs]) if paragraphs else ''
            return title, content, time_published, url
        except Exception as e:
            print(f"Error parsing HTML for URL '{url}': {e}")
    return 'No title', '', datetime.now().strftime('%Y-%m-%d %H:%M:%S'), url

# Function to fetch and parse RSS feeds based on keywords
def fetch_rss_feed(keyword):
    rss_url = f"https://news.google.com/rss/search?q={keyword}"
    try:
        feed = feedparser.parse(rss_url)
        articles = []
        for entry in feed.entries:
            title = entry.title
            content = entry.summary if 'summary' in entry else ''
            time_published = parse_date(entry.published) if 'published' in entry else datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            url = entry.link
            articles.append({'Time': time_published, 'Keyword': keyword, 'Title': title, 'URL': url, 'Content': content})
        return articles
    except Exception as e:
        print(f"Error fetching RSS feed for keyword '{keyword}': {e}")
        return []

# Keywords to search
keywords = ["dementia", "manastik"]

# Dictionary to hold DataFrames for each keyword
dataframes = {}

for keyword in keywords:
    data = []

    # Fetch articles from web search
    links = get_articles_from_web(keyword, num_results=200)
    for link in links:
        title, content, time_published, url = get_article_details_from_web(link)
        data.append({'Time': time_published, 'Keyword': keyword, 'Title': title, 'URL': url, 'Content': content})

    # Fetch articles from RSS feeds
    rss_articles = fetch_rss_feed(keyword)
    data.extend(rss_articles)

    # Create a DataFrame for each keyword
    df = pd.DataFrame(data)

    if not df.empty:
        df['Time'] = pd.to_datetime(df['Time'])
        df.sort_values(by='Time', ascending=True, inplace=True)
        dataframes[keyword] = df
    else:
        print(f"No data found for keyword: {keyword}")

# Save DataFrames to different sheets in a single Excel file
if dataframes:
    output_file = 'Keyword_scrapingresult.xlsx'
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        for keyword, df in dataframes.items():
            df.to_excel(writer, sheet_name=keyword, index=False)
            worksheet = writer.sheets[keyword]
            
            # Adjust column width for clarity
            for i, col in enumerate(df.columns):
                max_len = df[col].astype(str).map(len).max()
                max_len = max(max_len, len(col))
                worksheet.set_column(i, i, max_len + 2)
                
            # Format date columns
            date_format = writer.book.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})
            worksheet.set_column('A:A', 20, date_format)
    print(f"Data saved to {output_file}")
else:
    print("No data to save.")
    
    def add_numbers(a, b):
    
       return a + b

# Example usage:
result = add_numbers(5, 3)
print("The sum is:", result)

