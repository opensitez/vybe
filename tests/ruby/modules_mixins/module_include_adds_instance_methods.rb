# vybe-test: ruby/modules_mixins/module_include_adds_instance_methods
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module Greetable
  def greet; "Hello, I am #{name}"; end
end
class Person
  include Greetable
  attr_reader :name
  def initialize(n); @name = n; end
end
puts Person.new("Alice").greet
