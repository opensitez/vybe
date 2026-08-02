# vybe-test: python/concurrency_runtime/multiprocessing_spawn
# origin: languages/python/tests/python/test_concurrency_runtime.rs
# vybe-test-mode: compile

import multiprocessing
multiprocessing.get_context('spawn')
