# vybe-test: python/generators_core/generator_in_for_loop
# origin: languages/python/tests/python/test_generators_core.rs

def g():
 yield 'a'
 yield 'b'
out = ''
for ch in g():
 out += ch
print(out)
