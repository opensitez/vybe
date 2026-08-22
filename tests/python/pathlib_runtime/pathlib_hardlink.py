# vybe-test: python/pathlib_runtime/pathlib_hardlink
# origin: languages/python/tests/python/test_pathlib_runtime.rs

from pathlib import Path
hasattr(Path('.'), 'hardlink_to')
