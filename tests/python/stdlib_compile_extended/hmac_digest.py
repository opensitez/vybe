# vybe-test: python/stdlib_compile_extended/hmac_digest
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

import hmac
hmac.new(b'key', b'msg', 'sha256')
