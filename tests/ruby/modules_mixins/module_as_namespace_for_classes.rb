# vybe-test: ruby/modules_mixins/module_as_namespace_for_classes
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module Animals
  class Dog
    def speak; "woof"; end
  end
  class Cat
    def speak; "meow"; end
  end
end
puts Animals::Dog.new.speak
puts Animals::Cat.new.speak
