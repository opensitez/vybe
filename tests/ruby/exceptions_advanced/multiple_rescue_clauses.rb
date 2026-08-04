# vybe-test: ruby/exceptions_advanced/multiple_rescue_clauses
# origin: languages/ruby/tests/ruby/test_exceptions_advanced.rs
# vybe-test-mode: compile

class FooError < StandardError
end
class BarError < StandardError
end
begin
  raise FooError
rescue FooError
  puts 'foo'
rescue BarError
  puts 'bar'
end
