import configparser
from datetime import datetime
import re
import smtplib
import time
from email.mime.text import MIMEText

import openpyxl
import undetected_chromedriver as uc
from bs4 import BeautifulSoup


def create_driver(proxy_server):
    """配置并返回一个 Chrome WebDriver 实例。"""
    options = uc.ChromeOptions()
    if proxy_server:
        options.add_argument(f'--proxy-server={proxy_server}')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    return uc.Chrome(options=options)

def scrape_job_links(driver, base_url, keyword):
    """从主搜索页面抓取所有独立的职位链接。"""
    search_url = f"{base_url}?Keyword={keyword}"
    driver.get(search_url)
    print(f"Successfully opened: {driver.title}")

    soup = BeautifulSoup(driver.page_source, 'html.parser')
    total_pages = 1
    pagination_info = soup.find('li', class_='space')
    if pagination_info:
        match = re.search(r'Total (\d+) Page\(s\)', pagination_info.get_text())
        if match:
            total_pages = int(match.group(1))
    print(f"Found {total_pages} pages of job listings.")

    all_links = set()
    link_pattern = re.compile(r'^https://jump\.mingpao\.com/job/detail/Jobs/2')
    for page_num in range(1, total_pages + 1):
        page_url = f"{search_url}&Page={page_num}"
        print(f"Scraping job links from page: {page_url}")
        driver.get(page_url)
        time.sleep(2)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        for link in soup.find_all('a', href=link_pattern):
            all_links.add(link.get('href'))

    print(f"Found {len(all_links)} unique job links.")
    return list(all_links)


def extract_emails_from_html(soup):
    # 只查找指定路径下的内容
    target_div = soup.select_one('div.margin1em0.pull-left')
    if not target_div:
        return []

    # 方法 1：提取 href 中的 mailto
    mailto_links = target_div.select('a[href^=mailto]')
    emails = []
    for link in mailto_links:
        href = link.get('href')
        if href:
            match = re.search(r"mailto:([^?]+)", href)
            if match:
                emails.append(match.group(1))

    # 方法 2（可选）：兜底再从纯文本中正则提取
    # email_pattern = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
    # emails.extend(email_pattern.findall(target_div.get_text()))

    # 去重
    return list(set(emails))


def scrape_job_details(driver, job_links):
    """访问每个职位链接并提取职位名称和电子邮件。"""
    details = []
    email_pattern = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")

    for link in job_links:
        print(f"Scraping details from: {link}")
        driver.get(link)
        time.sleep(1)
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # 提取职位名称
        position_tag = soup.select_one('.color_position.txt_16px.bold h1.h3')
        position = position_tag.get_text(strip=True) if position_tag else "Not Found"

        # 提取邮件地址
        emails = extract_emails_from_html(soup)
        details.append({
            'link': link,
            'position': position,
            'emails': ", ".join(emails) if emails else "Not Found"
        })
    return details

def save_to_excel(details, filename):
    """将提取的职位详情保存到Excel文件。"""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Job Details"
    sheet.append(['Link', 'Position', 'Emails'])

    for item in details:
        sheet.append([item['link'], item['position'], item['emails']])

    workbook.save(filename)
    print(f"Saved {len(details)} job details to {filename}")

def send_email(subject, body, config):
    """发送邮件通知。"""
    if not config.getboolean('Email', 'enabled', fallback=False):
        print("Email sending is disabled.")
        return

    try:
        msg = MIMEText(body, 'plain', 'utf-8')
        msg['Subject'] = subject
        msg['From'] = config['Email']['sender_email']
        msg['To'] = config['Email']['recipient_email']

        with smtplib.SMTP(config['Email']['smtp_server'], config.getint('Email', 'smtp_port')) as server:
            server.starttls()
            server.login(config['Email']['sender_email'], config['Email']['sender_password'])
            server.send_message(msg)
        print("Email notification sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")

def main():
    """主程序流程控制器。"""
    config = configparser.ConfigParser()
    config.read('config.ini', encoding='utf-8')

    driver = create_driver(config.get('Scraper', 'proxy', fallback=None))
    try:
        job_links = scrape_job_links(driver, config['Scraper']['base_url'], config['Scraper']['keyword'])
        if job_links:
            details = scrape_job_details(driver, job_links)
            output_file = config['Scraper']['keyword'] + str(datetime.now().timestamp() * 1000) +config['Scraper']['output_file']
            save_to_excel(details, output_file)

            email_subject = f"JUMP Job Scraping Report - {len(details)} jobs found"
            email_body = f"Finished scraping.\n\nFound {len(details)} jobs. See the Excel file '{output_file}' for details."
            send_email(email_subject, email_body, config)
        else:
            print("No job links were found to scrape details from.")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
