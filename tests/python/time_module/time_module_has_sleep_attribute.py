# vybe-test: python/time_module/time_module_has_sleep_attribute
# origin: languages/python/tests/python/test_time_module.rs

hasattr(__import__('time'), 'sleep')
