# vybe-test: ruby/exceptions_advanced/rescue_exception_catches_all
# origin: languages/ruby/tests/ruby/test_exceptions_advanced.rs
# vybe-test-mode: compile

begin
  raise 'anything'
rescue Exception
  puts 'caught all'
end
