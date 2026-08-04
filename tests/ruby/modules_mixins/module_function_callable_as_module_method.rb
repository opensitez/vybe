# vybe-test: ruby/modules_mixins/module_function_callable_as_module_method
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module MathUtils
  module_function
  def square(x); x * x; end
  def cube(x); x ** 3; end
end
puts MathUtils.square(4)
puts MathUtils.cube(3)
