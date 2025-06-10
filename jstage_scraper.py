import json
import logging
import os
import time
from urllib.parse import urljoin
import asyncio

import pandas as pd
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright


class JstageScraper:
    """A web scraper for J-STAGE"""
    def __init__(self, output_json='jstage_results.json', output_excel='jstage_results.xlsx', max_runtime_minutes=15):
        """Initialize the scraper"""
        # Set up logging
        self.logger = logging.getLogger('jstage_scraper')
        self.logger.setLevel(logging.DEBUG)
        
        # Create console handler with a higher log level
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.DEBUG)
        
        # Create file handler which logs even debug messages
        file_handler = logging.FileHandler('jstage_scraper.log')
        file_handler.setLevel(logging.DEBUG)
        
        # Create formatters and add them to the handlers
        console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        console_handler.setFormatter(console_formatter)
        file_handler.setFormatter(file_formatter)
        
        # Add the handlers to the logger
        if not self.logger.handlers:
            self.logger.addHandler(console_handler)
            self.logger.addHandler(file_handler)

        self.base_url = "https://www.jstage.jst.go.jp"
        self.output_json = output_json
        self.output_excel = output_excel
        self.max_runtime_seconds = max_runtime_minutes * 60
        self.start_time = None
        self.visited_urls = set()
        self.results = []
        
        try:
            self.init_output_files()
        except Exception as e:
            self.logger.error(f"Error initializing output files: {e}")
            
    def init_output_files(self):
        """Initialize or clear the output files."""
        # For JSON
        with open(self.output_json, 'w') as f:
            json.dump([], f)
        # For Excel
        pd.DataFrame([]).to_excel(self.output_excel, index=False)
        self.logger.info("Output files initialized successfully")

    async def fetch_issue_urls(self, journal_url):
        self.logger.info(f"Fetching issues from: {journal_url}")
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            page = await browser.new_page()
            await page.goto(journal_url, timeout=60000)
            await page.wait_for_timeout(2000)
            volume_options = await page.eval_on_selector_all(
                'select#selectVol option',
                'options => options.map(o => o.value).filter(v => v && v !== "")'
            )
            for vol_val in volume_options:
                await page.select_option('select#selectVol', vol_val)
                for _ in range(20):
                    try:
                        await page.wait_for_selector('select#selectIssue:not([disabled])', timeout=500)
                        issue_options = await page.eval_on_selector_all(
                            'select#selectIssue option',
                            'options => options.map(o => o.value).filter(v => v && v !== "")'
                        )
                        if len(issue_options) > 0:
                            break
                    except Exception:
                        pass
                    await page.wait_for_timeout(500)
                else:
                    self.logger.warning(f"No issues found for volume {vol_val}")
                    continue
                issue_options = await page.eval_on_selector_all(
                    'select#selectIssue option',
                    'options => options.map(o => o.value).filter(v => v && v !== "")'
                )
                for iss_val in issue_options:
                    try:
                        await page.select_option('select#selectIssue', iss_val, timeout=3000)
                    except Exception:
                        await page.evaluate(f'''
                            () => {{
                                const sel = document.querySelector('select#selectIssue');
                                sel.value = "{iss_val}";
                                sel.dispatchEvent(new Event('change', {{ bubbles: true }}));
                            }}
                        ''')
                        await page.wait_for_timeout(500)
                    await page.wait_for_timeout(1000)
                    async with page.expect_navigation():
                        await page.click('input#goBtn')
                    issue_url = page.url
                    self.logger.info(f"Found issue: {issue_url}")
                    yield issue_url
                    await page.goto(journal_url, timeout=60000)
                    await page.wait_for_timeout(1000)
            await browser.close()

    async def fetch_page(self, url, max_retries=3):
        """Fetch a page with retries using Playwright"""
        for attempt in range(max_retries):
            try:
                async with async_playwright() as p:
                    browser = await p.chromium.launch(headless=True)
                    page = await browser.new_page(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
                    await page.goto(url, timeout=60000)
                    try:
                        await page.wait_for_selector('div.global-article-title', timeout=15000)
                    except Exception:
                        self.logger.warning(f"Selector 'div.global-article-title' not found on {url}")
                    await page.wait_for_timeout(2000)
                    content = await page.content()
                    await browser.close()
                    return content
            except Exception as e:
                self.logger.warning(f"Error fetching {url} on attempt {attempt + 1}: {e}")
                if attempt < max_retries - 1:
                    await asyncio.sleep(5)
                else:
                    self.logger.error(f"Failed to fetch {url} after {max_retries} attempts.")
                    return None

    async def get_articles_from_issue(self, issue_url):
        """Get articles from an issue page"""
        self.logger.info(f"Fetching articles from issue: {issue_url}")
        html = await self.fetch_page(issue_url)
        if not html:
            return []
        soup = BeautifulSoup(html, 'html.parser')
        articles = set()
        # Only collect links that match the full J-STAGE article URL pattern
        for a in soup.find_all('a', href=True):
            href = a['href']
            if '/article/' in href and '_article/-char/en' in href:
                full_url = urljoin(self.base_url, href)
                articles.add(full_url)
        self.logger.info(f"Found {len(articles)} articles in issue.")
        return list(articles)

    async def get_article_metadata(self, article_url, journal_name, subject_area):
        """Extract metadata from an article page"""
        self.logger.info(f"Extracting metadata from article: {article_url}")
        html = await self.fetch_page(article_url)
        if not html:
            return None
            
        soup = BeautifulSoup(html, 'html.parser')
        try:
            title_elem = soup.find('div', class_='global-article-title')
            title = title_elem.get_text(strip=True) if title_elem else ''
            authors = []
            for author in soup.find_all('span', class_='author-name'):
                authors.append(author.get_text(strip=True))
            doi = ''
            doi_elem = soup.find('a', href=True, string=lambda s: s and s.startswith('https://doi.org/'))
            if doi_elem:
                doi = doi_elem.get_text(strip=True)
            abstract = ''
            abstract_elem = soup.find('div', class_='abstract')
            if abstract_elem:
                abstract = abstract_elem.get_text(strip=True)
            keywords = []
            keywords_elem = soup.find('div', class_='keywords')
            if keywords_elem:
                for kw in keywords_elem.find_all('span', class_='keyword'):
                    keywords.append(kw.get_text(strip=True))
            pdf_url = ''
            pdf_link = soup.find('a', href=True, string=lambda s: s and 'PDF' in s)
            if pdf_link:
                pdf_url = urljoin(self.base_url, pdf_link['href'])
            return {
                'subject_area': subject_area,
                'journal_name': journal_name,
                'title': title,
                'url': article_url,
                'authors': authors,
                'doi': doi,
                'abstract': abstract,
                'keywords': keywords,
                'pdf_url': pdf_url
            }
        except Exception as e:
            self.logger.error(f"Error extracting metadata from {article_url}: {e}")
            return None

    async def run(self):
        self.start_time = time.time()
        self.logger.info("\nStarting J-STAGE scraper...")
        self.logger.info(f"Scraper will run for {self.max_runtime_seconds / 60} minutes maximum.")
        math_journals = [
            {'name': 'Tohoku Mathematical Journal', 'url': 'https://www.jstage.jst.go.jp/browse/tmj'},
            {'name': 'Journal of the Mathematical Society of Japan', 'url': 'https://www.jstage.jst.go.jp/browse/jmsj'},
            {'name': 'Kodai Mathematical Journal', 'url': 'https://www.jstage.jst.go.jp/browse/kmj'}
        ]
        subject_area = 'Mathematics'
        for journal in math_journals:
            if time.time() - self.start_time > self.max_runtime_seconds:
                self.logger.info("Maximum runtime reached. Stopping scraper.")
                break
            self.logger.info(f"\nProcessing journal: {journal['name']}")
            async for issue_url in self.fetch_issue_urls(journal['url']):
                if time.time() - self.start_time > self.max_runtime_seconds:
                    self.logger.info("Maximum runtime reached. Stopping scraper.")
                    break
                article_links = await self.get_articles_from_issue(issue_url)
                for article_url in article_links:
                    if article_url in self.visited_urls:
                        continue
                    self.visited_urls.add(article_url)
                    data = await self.get_article_metadata(article_url, journal['name'], subject_area)
                    if data:
                        self.results.append(data)
                        self.save_results()
        self.logger.info("\nScraping completed!")
        self.log_summary()

    def save_results(self):
        """Save the scraped results to JSON and Excel"""
        if not self.results:
            self.logger.warning("No results to save")
            return
            
        try:
            # Save to JSON
            with open(self.output_json, 'w') as f:
                json.dump(self.results, f, indent=4)
            
            # Save to Excel
            df = pd.DataFrame(self.results)
            df.to_excel(self.output_excel, index=False)
        except Exception as e:
            self.logger.error(f"Error saving results: {e}")

    def log_summary(self):
        """Log a summary of the scraping session"""
        end_time = time.time()
        total_time = end_time - self.start_time
        self.logger.info(f"Total articles processed: {len(self.results)}")
        self.logger.info(f"Total unique URLs visited: {len(self.visited_urls)}")
        self.logger.info(f"Total time elapsed: {total_time:.2f} seconds")

if __name__ == '__main__':
    import sys
    
    in_background = '--background' in sys.argv
    
    if in_background:
        print("Starting scraper in background mode...")
        # Detach from the console
        if os.name == 'nt':
            import subprocess
            subprocess.Popen([sys.executable] + [arg for arg in sys.argv if arg != '--background'], creationflags=subprocess.CREATE_NO_WINDOW)
        else:
            if os.fork():
                sys.exit()

    scraper = JstageScraper()
    
    print(f"Results will be saved to {scraper.output_json} and {scraper.output_excel}")
    print(f"Maximum runtime: {scraper.max_runtime_seconds / 60} minutes")
    print("Press Ctrl+C to stop the scraper at any time")

    try:
        asyncio.run(scraper.run())
    except KeyboardInterrupt:
        print("\nScraper interrupted by user.")
    finally:
        scraper.save_results()
        print("Scraper has finished.")