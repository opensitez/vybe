# vybe-test: python/for_else_core/for_else_generator_exhausted
# origin: languages/python/tests/python/test_for_else_core.rs

def gen():
 yield 1
 yield 2
for x in gen():
 pass
else:
 print('gen')
