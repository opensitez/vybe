# vybe-test: python/diagnostics_runtime/warnings_warn_stacklevel
# origin: languages/python/tests/python/test_diagnostics_runtime.rs
# vybe-test-mode: compile

import warnings
warnings.warn('m', stacklevel=2)
