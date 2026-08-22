# vybe-test: python/module_introspection_extended/runpy_run_path
# origin: languages/python/tests/python/test_module_introspection_extended.rs
# `run_path('.')` needs a directory containing `__main__.py` — the repo
# root has none. Point it at a real script.
import os, runpy, tempfile
_d = tempfile.mkdtemp()
_s = os.path.join(_d, 'script.py')
open(_s, 'w').write('value = 1\n')
runpy.run_path(_s, run_name='__main__')
