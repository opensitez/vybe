# vybe-test: ruby/exceptions_advanced/rescue_in_method_body
# origin: languages/ruby/tests/ruby/test_exceptions_advanced.rs
# vybe-test-mode: compile

def safe_divide(a, b)
  a / b
rescue ZeroDivisionError
  0
end
safe_divide(10, 0)
