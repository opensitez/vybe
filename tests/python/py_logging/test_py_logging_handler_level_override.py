# vybe-test: python/py_logging/test_py_logging_handler_level_override
# origin: languages/python/tests/python/test_py_logging.rs

import logging, io

buf = io.StringIO()
handler = logging.StreamHandler(buf)
handler.setLevel(logging.WARNING)

logger = logging.getLogger("filtered_logger")
logger.setLevel(logging.DEBUG)
logger.addHandler(handler)
logger.propagate = False

logger.debug("debug message")   # filtered out by handler
logger.warning("warning message")  # passes

lines = [l.strip() for l in buf.getvalue().strip().splitlines() if l.strip()]
print(len(lines))
print("warning message" in lines[0])
