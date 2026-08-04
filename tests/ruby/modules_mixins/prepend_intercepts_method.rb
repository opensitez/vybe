# vybe-test: ruby/modules_mixins/prepend_intercepts_method
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module Logging
  def compute(x)
    result = super
    result
  end
end
class Calculator
  prepend Logging
  def compute(x); x * 2; end
end
puts Calculator.new.compute(5)
