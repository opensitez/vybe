# vybe-test: ruby/modules_mixins/nested_module_constant
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module Outer
  module Inner
    VALUE = 42
  end
end
puts Outer::Inner::VALUE
