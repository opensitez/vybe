# vybe-test: ruby/modules_mixins/extend_on_specific_instance
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module Serializable
  def serialize; "data"; end
end
obj1 = Object.new
obj2 = Object.new
obj1.extend(Serializable)
puts obj1.respond_to?(:serialize)
puts obj2.respond_to?(:serialize)
