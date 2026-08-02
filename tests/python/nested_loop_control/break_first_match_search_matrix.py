# vybe-test: python/nested_loop_control/break_first_match_search_matrix
# origin: languages/python/tests/python/test_nested_loop_control.rs

grid = [[0, 0], [0, 5], [0, 0]]
found = None
for r in range(3):
 for c in range(2):
  if grid[r][c] == 5:
   found = (r, c)
   break
 if found:
  break
print(found)
