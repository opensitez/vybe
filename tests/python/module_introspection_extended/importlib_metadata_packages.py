# vybe-test: python/module_introspection_extended/importlib_metadata_packages
# origin: languages/python/tests/python/test_module_introspection_extended.rs

try:
 import importlib.metadata as md
 print(callable(md.packages_distributions))
except ImportError:
 print('skip')
