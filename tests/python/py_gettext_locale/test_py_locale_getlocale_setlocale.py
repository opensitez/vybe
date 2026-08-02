# vybe-test: python/py_gettext_locale/test_py_locale_getlocale_setlocale
# origin: languages/python/tests/python/test_py_gettext_locale.rs

import locale

current = locale.getlocale()
print(isinstance(current, tuple))

try:
    locale.setlocale(locale.LC_ALL, "C")
    print(locale.getlocale())
except locale.Error:
    print("Locale setting failed")
