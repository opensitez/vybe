# vybe-test: ruby/functions/lambda_call
# origin: languages/ruby/tests/ruby/test_functions.rs
# vybe-test-mode: compile

f = ->(x) { x * 2 }
puts f.call(5)
