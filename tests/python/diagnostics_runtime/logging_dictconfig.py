# vybe-test: python/diagnostics_runtime/logging_dictconfig
# origin: languages/python/tests/python/test_diagnostics_runtime.rs
# `dictConfig` REQUIRES a 'version' key — `{}` raises
# "dictionary doesn't specify a version".
import logging.config
logging.config.dictConfig({'version': 1})
