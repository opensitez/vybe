# vybe-test: python/function_signatures_spec/from_import_alias_parenthesized_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs
# The package this import names never existed. Build a REAL one so the
# import FORM under test actually resolves.
import os, sys, tempfile
_d = tempfile.mkdtemp()
os.makedirs(os.path.join(_d, 'package', 'subpackage'))
open(os.path.join(_d, 'package', '__init__.py'), 'w').close()
open(os.path.join(_d, 'package', 'subpackage', '__init__.py'), 'w').close()
open(os.path.join(_d, 'package', 'module.py'), 'w').write('name = 1\nother = 2\n')
open(os.path.join(_d, 'package', 'subpackage', 'tool.py'), 'w').write('v = 1\n')
open(os.path.join(_d, 'pkg.py'), 'w').write('v = 3\n')
sys.path.insert(0, _d)

from package.module import (
    name as alias,
    other
)
