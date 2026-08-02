# vybe-test: python/python_http_cookies_morsel/test_http_cookies_output_sep_header
# origin: languages/python/tests/python/test_python_http_cookies_morsel.rs

from http.cookies import SimpleCookie
c = SimpleCookie()
c["a"] = "1"
c["b"] = "2"
print(c.output(sep="; "))
