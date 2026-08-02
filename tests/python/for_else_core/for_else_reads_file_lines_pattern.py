# vybe-test: python/for_else_core/for_else_reads_file_lines_pattern
# origin: languages/python/tests/python/test_for_else_core.rs

lines = ['ok', 'done']
for line in lines:
 if line == 'error':
  print('fail')
  break
else:
 print('pass')
