# vybe-test: python/filesystem_utils/glob_recursive
# origin: languages/python/tests/python/test_filesystem_utils.rs
# vybe-test-mode: compile

import glob
glob.glob('**/*', recursive=True)
