# vybe-test: ruby/classes/module_include
# origin: languages/ruby/tests/ruby/test_classes.rs
# vybe-test-mode: compile

module Greetable
  def greet
    puts 'hello'
  end
end
class Person
  include Greetable
end
