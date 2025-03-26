import requests
import re
import time
import os
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
from concurrent.futures import ThreadPoolExecutor

class EmailCrawler:
    def __init__(self, base_url, max_depth=2, max_threads=10, output_file="emails.xlsx"):
        self.base_url = base_url
        self.max_depth = max_depth
        self.max_threads = max_threads
        self.output_file = output_file
        self.visited_urls = set()  # Tracks scanned pages
        self.emails = set()  # Stores found emails
        self.executor = ThreadPoolExecutor(max_workers=max_threads)  # Multi-threading

        # Load existing emails if the file exists
        self.existing_emails = self.load_existing_emails()

    def load_existing_emails(self):
        """Load existing emails from the Excel file to avoid duplicates"""
        if os.path.exists(self.output_file):
            try:
                df = pd.read_excel(self.output_file)
                return set(df["Email"].dropna().unique())  # Convert existing emails to a set
            except Exception as e:
                print(f"⚠️ Error loading existing emails: {e}")
                return set()
        return set()

    def fetch_page(self, url):
        """Fetch the page content"""
        try:
            headers = {"User-Agent": "Mozilla/5.0"}  # Prevent getting blocked
            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code == 200:
                return response.text
        except requests.RequestException:
            return None
        return None

    def extract_emails(self, text):
        """Extract emails using regex and remove duplicates"""
        new_emails = set(re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text))
        new_emails -= self.existing_emails  # Skip already saved emails
        self.emails.update(new_emails)

    def extract_links(self, url, html):
        """Extracts all internal links from a page"""
        soup = BeautifulSoup(html, "html.parser")
        links = set()
        for a_tag in soup.find_all("a", href=True):
            absolute_url = urljoin(url, a_tag["href"])
            if self.is_valid_url(absolute_url):
                links.add(absolute_url)
        return links

    def is_valid_url(self, url):
        """Checks if a URL is within the same domain and not visited"""
        parsed_url = urlparse(url)
        return (
            parsed_url.netloc == urlparse(self.base_url).netloc and
            url not in self.visited_urls
        )

    def crawl(self, url, depth=0):
        """Recursively crawls the website using multi-threading"""
        if depth > self.max_depth or url in self.visited_urls:
            return  # Stop if max depth is reached or URL is already scanned

        print(f"Crawling: {url}")
        self.visited_urls.add(url)  # Mark as visited

        html = self.fetch_page(url)
        if not html:
            return

        self.extract_emails(html)  # Extract emails from page
        links = self.extract_links(url, html)

        # Submit multiple requests in parallel
        for link in links:
            self.executor.submit(self.crawl, link, depth + 1)
            time.sleep(0.5)  # Prevent overloading the server

    def save_to_excel(self):
        """Save emails to an Excel file, adding new emails without duplicates"""
        if not self.emails:
            print("❌ No new emails found!")
            return

        # Create a DataFrame with Domain & Email
        new_data = pd.DataFrame({"Domain": [self.base_url] * len(self.emails), "Email": list(self.emails)})

        if os.path.exists(self.output_file):
            # Load existing data and append new rows
            df_existing = pd.read_excel(self.output_file)
            df_final = pd.concat([df_existing, new_data]).drop_duplicates(subset="Email", keep="first")
        else:
            df_final = new_data

        # Save updated data
        df_final.to_excel(self.output_file, index=False)
        print(f"✅ Emails saved to {self.output_file} (Total: {len(df_final)} emails)")

    def run(self):
        """Starts the crawler"""
        self.crawl(self.base_url)
        self.executor.shutdown(wait=True)  # Ensure all threads finish
        self.save_to_excel()  # Save emails to file

# Get domain from user input
if __name__ == "__main__":
    domain = input("Enter the website URL (e.g., https://example.com): ").strip()
    if not domain.startswith("http"):
        domain = "https://" + domain  # Ensure proper URL format

    crawler = EmailCrawler(domain, max_depth=2, max_threads=10)
    crawler.run()
