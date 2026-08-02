# vybe-test: python/generators_core/generator_read_file_lines_style
# origin: languages/python/tests/python/test_generators_core.rs

def lines():
 for s in ['a', 'b']:
  yield s.upper()
print(list(lines()))
