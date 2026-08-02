# vybe-test: python/concurrency_runtime/subprocess_run_echo
# origin: languages/python/tests/python/test_concurrency_runtime.rs
# vybe-test-mode: compile

import subprocess
subprocess.run(['echo', 'hi'], capture_output=True)
