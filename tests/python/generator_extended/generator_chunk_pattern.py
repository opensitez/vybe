# vybe-test: python/generator_extended/generator_chunk_pattern
# origin: languages/python/tests/python/test_generator_extended.rs

def chunks(it, n):
 buf = []
 for x in it:
  buf.append(x)
  if len(buf) == n:
   yield buf
   buf = []
print(list(chunks([1,2,3,4], 2)))
