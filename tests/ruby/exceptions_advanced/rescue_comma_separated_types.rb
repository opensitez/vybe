# vybe-test: ruby/exceptions_advanced/rescue_comma_separated_types
# origin: languages/ruby/tests/ruby/test_exceptions_advanced.rs
# vybe-test-mode: compile

class FooError < StandardError
end
class BarError < StandardError
end
begin
  raise FooError
rescue FooError, BarError
  puts 'caught one of them'
end
