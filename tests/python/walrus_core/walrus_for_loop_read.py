# vybe-test: python/walrus_core/walrus_for_loop_read
# origin: languages/python/tests/python/test_walrus_core.rs

pairs = [('a', 1)]
for k, v in pairs:
 if (label := k + str(v)):
  print(label)
