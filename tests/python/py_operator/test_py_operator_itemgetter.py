# vybe-test: python/py_operator/test_py_operator_itemgetter
# origin: languages/python/tests/python/test_py_operator.rs

from operator import itemgetter

data = [{"name": "Bob", "age": 25}, {"name": "Alice", "age": 30}, {"name": "Charlie", "age": 20}]
get_name = itemgetter("name")
print(get_name(data[0]))

sorted_by_age = sorted(data, key=itemgetter("age"))
print([d["name"] for d in sorted_by_age])

# Multiple keys
get_both = itemgetter("name", "age")
print(get_both(data[0]))
