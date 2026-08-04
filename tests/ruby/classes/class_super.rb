# vybe-test: ruby/classes/class_super
# origin: languages/ruby/tests/ruby/test_classes.rs
# vybe-test-mode: compile

class Animal
  def speak
    'generic'
  end
end
class Dog < Animal
  def speak
    super
  end
end
