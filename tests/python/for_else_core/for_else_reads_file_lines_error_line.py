# vybe-test: python/for_else_core/for_else_reads_file_lines_error_line
# origin: languages/python/tests/python/test_for_else_core.rs

lines = ['ok', 'error']
for line in lines:
 if line == 'error':
  print('fail')
  break
else:
 print('pass')
