# vybe-test: ruby/classes_advanced/frozen_check_unfrozen
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Box
  attr_accessor :value
end
b = Box.new
b.frozen?
