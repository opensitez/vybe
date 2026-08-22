# vybe-test: python/csv_markup_runtime/html_parser_feed
# origin: languages/python/tests/python/test_csv_markup_runtime.rs

from html.parser import HTMLParser
p = HTMLParser()
p.feed('<p>')
