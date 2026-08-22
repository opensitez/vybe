# vybe-test: python/module_introspection_extended/module_spec_from_loader
# origin: languages/python/tests/python/test_module_introspection_extended.rs

import importlib.util
import json
importlib.util.spec_from_loader('x', json.__loader__)
