# vybe-test: python/bytes_decode_encode/bytes_iterable_in_for_loop_sum
# origin: languages/python/tests/python/test_bytes_decode_encode.rs

total = 0
for b in b'\x01\x02\x03':
 total += b
print(total)
