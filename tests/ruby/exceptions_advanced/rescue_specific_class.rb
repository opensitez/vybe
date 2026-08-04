# vybe-test: ruby/exceptions_advanced/rescue_specific_class
# origin: languages/ruby/tests/ruby/test_exceptions_advanced.rs
# vybe-test-mode: compile

class NetworkError < StandardError
end
begin
  raise NetworkError
rescue NetworkError
  puts 'caught network error'
end
