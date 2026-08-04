# vybe-test: ruby/classes_advanced/class_eval_add_method
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Robot
end
Robot.class_eval do
  def beep
    'beep'
  end
end
Robot.new.beep
