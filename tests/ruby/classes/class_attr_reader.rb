# vybe-test: ruby/classes/class_attr_reader
# origin: languages/ruby/tests/ruby/test_classes.rs
# vybe-test-mode: compile

class Dog
  attr_reader :name
  def initialize(name)
    @name = name
  end
end
