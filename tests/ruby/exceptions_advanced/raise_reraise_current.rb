# vybe-test: ruby/exceptions_advanced/raise_reraise_current
# origin: languages/ruby/tests/ruby/test_exceptions_advanced.rs
# vybe-test-mode: compile

def risky
  raise 'original'
rescue => e
  raise
end
risky rescue nil
