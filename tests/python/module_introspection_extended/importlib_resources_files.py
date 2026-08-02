# vybe-test: python/module_introspection_extended/importlib_resources_files
# origin: languages/python/tests/python/test_module_introspection_extended.rs
# vybe-test-mode: compile

from importlib import resources
resources.files('json')
