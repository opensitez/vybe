# vybe-test: ruby/modules_mixins/module_extend_adds_class_methods
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module ClassMethods
  def create(name); new(name); end
end
class Robot
  extend ClassMethods
  attr_reader :name
  def initialize(n); @name = n; end
end
puts Robot.create("R2D2").name
