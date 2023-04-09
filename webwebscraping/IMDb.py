import scrapy
# estrutura base
"""class QuotsSpider(scrapy.Spider):
    name = ''
    start_url = []

    def parse(self, response):
        pass"""

class QuotsSpider(scrapy.Spider):
    name = 'QuotsSpider'
    start_urls = ['https://quotes.toscrape.com/']

    def parse(self, response):
        qoute = response.css('.quote::text')
        for q in qoute:
            print(q, end=', ')