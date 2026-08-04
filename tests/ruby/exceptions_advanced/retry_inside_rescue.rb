# vybe-test: ruby/exceptions_advanced/retry_inside_rescue
# origin: languages/ruby/tests/ruby/test_exceptions_advanced.rs
# vybe-test-mode: compile

attempts = 0
begin
  attempts += 1
  raise 'fail' if attempts < 3
rescue
  retry if attempts < 3
end
