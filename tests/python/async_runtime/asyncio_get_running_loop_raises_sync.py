# vybe-test: python/async_runtime/asyncio_get_running_loop_raises_sync
# origin: languages/python/tests/python/test_async_runtime.rs

import asyncio
try:
 asyncio.get_running_loop()
 print('has')
except RuntimeError:
 print('none')
