# vybe-test: python/decorators_extended/decorator_list_comprehension
# origin: languages/python/tests/python/test_decorators_extended.rs

def deco(f):
 return f
def make():
 return [deco(lambda i=i: i) for i in range(3)]
print(make()[2]())
