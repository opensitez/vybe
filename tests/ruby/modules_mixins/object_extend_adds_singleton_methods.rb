# vybe-test: ruby/modules_mixins/object_extend_adds_singleton_methods
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module Debuggable
  def debug_info; self.class.to_s + ': ' + inspect; end
end
obj = Object.new
obj.extend(Debuggable)
puts obj.respond_to?(:debug_info)
