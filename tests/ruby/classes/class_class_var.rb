# vybe-test: ruby/classes/class_class_var
# origin: languages/ruby/tests/ruby/test_classes.rs
# vybe-test-mode: compile

class Counter
  @@count = 0
  def initialize
    @@count += 1
  end
end
