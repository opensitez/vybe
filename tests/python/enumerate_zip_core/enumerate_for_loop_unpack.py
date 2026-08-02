# vybe-test: python/enumerate_zip_core/enumerate_for_loop_unpack
# origin: languages/python/tests/python/test_enumerate_zip_core.rs

out = []
for i, v in enumerate(['x', 'y']):
 out.append(f'{i}:{v}')
print(out)
