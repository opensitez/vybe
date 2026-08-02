# vybe-test: python/py_logging/test_py_logging_memory_handler
# origin: languages/python/tests/python/test_py_logging.rs

import logging

buf_handler = logging.handlers.MemoryHandler if hasattr(logging, 'handlers') else None

import logging.handlers, io

buf = io.StringIO()
target = logging.StreamHandler(buf)
target.setFormatter(logging.Formatter("%(message)s"))

mem_handler = logging.handlers.MemoryHandler(capacity=5, target=target)
logger = logging.getLogger("mem_log")
logger.addHandler(mem_handler)
logger.propagate = False
logger.setLevel(logging.DEBUG)

for i in range(3):
    logger.info(f"msg{i}")

mem_handler.flush()
lines = [l.strip() for l in buf.getvalue().strip().splitlines() if l.strip()]
print(lines)
