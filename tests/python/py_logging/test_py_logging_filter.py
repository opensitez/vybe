# vybe-test: python/py_logging/test_py_logging_filter
# origin: languages/python/tests/python/test_py_logging.rs

import logging, io

class PrefixFilter(logging.Filter):
    def __init__(self, prefix):
        self.prefix = prefix

    def filter(self, record):
        return record.getMessage().startswith(self.prefix)

buf = io.StringIO()
handler = logging.StreamHandler(buf)
handler.setFormatter(logging.Formatter("%(message)s"))
handler.addFilter(PrefixFilter("ALLOW:"))

logger = logging.getLogger("filtered")
logger.addHandler(handler)
logger.propagate = False
logger.setLevel(logging.DEBUG)

logger.info("ALLOW: this goes through")
logger.info("BLOCK: this is filtered")
logger.warning("ALLOW: so does this")

lines = [l.strip() for l in buf.getvalue().strip().splitlines() if l.strip()]
print(lines)
