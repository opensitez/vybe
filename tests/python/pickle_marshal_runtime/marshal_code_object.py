# vybe-test: python/pickle_marshal_runtime/marshal_code_object
# origin: languages/python/tests/python/test_pickle_marshal_runtime.rs
# vybe-test-mode: compile

import marshal
compile('1+1', '<s>', 'eval')
