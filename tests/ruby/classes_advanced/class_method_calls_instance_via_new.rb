# vybe-test: ruby/classes_advanced/class_method_calls_instance_via_new
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Builder
  def build
    'built'
  end
  def self.create_and_build
    Builder.new.build
  end
end
Builder.create_and_build
