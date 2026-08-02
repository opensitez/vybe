# vybe-test: python/slicing_extended/slice_object_index
# origin: languages/python/tests/python/test_slicing_extended.rs
# vybe-test-mode: compile

class S:
 def __getitem__(self, i):
  return i
S()[1:2]
