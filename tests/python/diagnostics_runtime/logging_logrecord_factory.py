# vybe-test: python/diagnostics_runtime/logging_logrecord_factory
# origin: languages/python/tests/python/test_diagnostics_runtime.rs
# vybe-test-mode: compile

import logging
logging.setLogRecordFactory(logging.LogRecord)
