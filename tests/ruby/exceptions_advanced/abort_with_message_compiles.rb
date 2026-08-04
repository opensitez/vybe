# vybe-test: ruby/exceptions_advanced/abort_with_message_compiles
# origin: languages/ruby/tests/ruby/test_exceptions_advanced.rs
# vybe-test-mode: compile

def maybe_abort(x)
  abort('fatal error') if x < 0
  x
end
