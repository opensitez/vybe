# vybe-test: ruby/classes_advanced/class_method_self_prefix
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class MathHelper
  def self.square(n)
    n * n
  end
end
result = MathHelper.square(7)
