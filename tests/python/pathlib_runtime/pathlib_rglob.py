# vybe-test: python/pathlib_runtime/pathlib_rglob
# origin: languages/python/tests/python/test_pathlib_runtime.rs

from pathlib import Path
list(Path('.').rglob('*.rs'))
