# vybe-test: ruby/classes_advanced/singleton_method_on_object
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

obj = Object.new
def obj.greet
  'hello from singleton'
end
obj.greet
