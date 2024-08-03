from bs4 import BeautifulSoup as BS
import requests
import openpyxl


def get_html(url):
    response = requests.get(url)
    if response.status_code == 200:
        return response.text
    return None

def get_post_links(html):
    soup = BS(html, 'html.parser')
    main_content = soup.find('div', class_='out__inner')
    posts1 = main_content.find('div', class_='news__row row row-flex')
    articles = posts1.find_all('div', class_='card')

    links = []
    for article in articles:
        link_tag = article.find('a', class_='card__link')
        if link_tag:
            link = link_tag.get('href')
            if link:
                if not link.startswith('https'):
                    link = 'https://24smi.org/article/' + link
                links.append(link)
    return links

def get_post_data(html):
    soup = BS(html, 'html.parser')
    main2 = soup.find('div', class_='category-news__main main col-8 col-md-12')
    main3 = soup.find('article', class_='article')
    title = main2.find('div', class_='category-news__head')
    date2 = title.find('div', class_='category-news__params')
    date1 = date2.find('div', class_='date date_light').text.strip()
    genre = date2.find('span', class_='badge').text.strip()
    title2 = main3.find('h1').text.strip()
    des = main3.find('p').text.strip()   
    # des = main3.find_all('p')
    # for d in des:
    #     print(d.text.strip())
    author = main2.find('div', class_='author-name').text.strip()
    img = main2.find('figure', class_='img')
    text_of_image = img.find('figcaption').text.strip()

    data = {
        'Name of article': title2,
        'Date of article': date1,
        'Category of article': genre,
        'Description of article': des,
        'Author of article': author,
        'Text of image inside in article': text_of_image
    }

    return data

def save_to_exel(data):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet['A1'] = 'Name of article'
    sheet['B1'] = 'Date of article'
    sheet['C1'] = 'Category of article'
    sheet['D1'] = 'Description of article'
    sheet['E1'] = 'Author of article'
    sheet['F1'] = 'Text of image inside in article'
    
    for i, item in enumerate(data, 2):
        sheet[f'A{i}'] = item['Name of article']
        sheet[f'B{i}'] = item['Date of article']
        sheet[f'C{i}'] = item['Category of article']
        sheet[f'D{i}'] = item['Description of article']
        sheet[f'E{i}'] = item['Author of article']
        sheet[f'F{i}'] = item['Text of image inside in article']
    
    workbook.save('project_data.xlsx')

def main():
    URL = 'https://24smi.org/article/'
    html = get_html(url=URL)
    if html:
        links = get_post_links(html=html)
        all_data = []
        for link in links:
            posts_links = get_html(url=link)
            if posts_links:
                post_data = get_post_data(html=posts_links)
                all_data.append(post_data)
        
        save_to_exel(all_data)

if __name__ == '__main__':
    main()