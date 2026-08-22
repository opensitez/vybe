# vybe-test: python/function_signatures_spec/import_dotted_as_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs
# `package.submodule` never existed. Build a REAL package so the import
# form under test actually resolves.
import os, sys, tempfile
_d = tempfile.mkdtemp()
os.makedirs(os.path.join(_d, 'package'))
open(os.path.join(_d, 'package', '__init__.py'), 'w').close()
open(os.path.join(_d, 'package', 'submodule.py'), 'w').write('value = 1\n')
open(os.path.join(_d, 'package', 'sibling.py'), 'w').write('value = 2\n')
sys.path.insert(0, _d)

import package.submodule as sm
