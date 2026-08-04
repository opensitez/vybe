# vybe-test: ruby/classes_advanced/attr_accessor_multiple
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Person
  attr_accessor :name, :age, :email
  def initialize(name, age, email)
    @name = name
    @age = age
    @email = email
  end
end
