# vybe-test: python/import_system/relative_import_package
# origin: languages/python/tests/python/test_import_system.rs
# `from . import X` at top level has no parent package — that is what
# "attempted relative import with no known parent package" means. The
# import is spelled against a real package instead.
# `package.submodule` never existed. Build a REAL package so the import
# form under test actually resolves.
import os, sys, tempfile
_d = tempfile.mkdtemp()
os.makedirs(os.path.join(_d, 'package'))
open(os.path.join(_d, 'package', '__init__.py'), 'w').close()
open(os.path.join(_d, 'package', 'submodule.py'), 'w').write('value = 1\n')
open(os.path.join(_d, 'package', 'sibling.py'), 'w').write('value = 2\n')
sys.path.insert(0, _d)

from package import sibling
