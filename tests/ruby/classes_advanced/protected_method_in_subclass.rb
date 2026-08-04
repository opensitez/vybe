# vybe-test: ruby/classes_advanced/protected_method_in_subclass
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Animal
  protected
  def secret
    'hidden'
  end
end
class Dog < Animal
  def reveal
    secret
  end
end
d = Dog.new
d.reveal
