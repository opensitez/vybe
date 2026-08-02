# vybe-test: python/python_operator_module/test_operator_itemgetter
# origin: languages/python/tests/python/test_python_operator_module.rs

import operator
data = [{'name': 'Bob', 'age': 30}, {'name': 'Alice', 'age': 25}]
by_age = sorted(data, key=operator.itemgetter('age'))
print([d['name'] for d in by_age])
