# vybe-test: python/filesystem_utils/tempfile_temporarydirectory_ctx
# origin: languages/python/tests/python/test_filesystem_utils.rs
# vybe-test-mode: compile

import tempfile
with tempfile.TemporaryDirectory() as d:
 pass
