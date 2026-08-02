# vybe-test: python/walrus_extended/walrus_in_class_body
# origin: languages/python/tests/python/test_walrus_extended.rs
# vybe-test-mode: compile

class C:
 x = (y := 1)
