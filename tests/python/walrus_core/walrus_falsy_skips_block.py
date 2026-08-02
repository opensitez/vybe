# vybe-test: python/walrus_core/walrus_falsy_skips_block
# origin: languages/python/tests/python/test_walrus_core.rs

items = ['']
if item := items[0]:
 print('yes')
else:
 print('no')
