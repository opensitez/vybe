# vybe-test: python/py_logging/test_py_logging_extra_contextual_fields
# origin: languages/python/tests/python/test_py_logging.rs

import logging, io

buf = io.StringIO()
handler = logging.StreamHandler(buf)
handler.setFormatter(logging.Formatter("%(user)s:%(message)s"))

logger = logging.getLogger("extra_test")
logger.addHandler(handler)
logger.propagate = False
logger.setLevel(logging.INFO)

logger.info("login attempt", extra={"user": "alice"})
logger.info("logout", extra={"user": "bob"})

lines = [l.strip() for l in buf.getvalue().strip().splitlines() if l.strip()]
print(lines)
