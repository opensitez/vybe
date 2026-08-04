# vybe-test: ruby/modules_mixins/included_callback_fires_on_include
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module Hookable
  def self.included(base)
    base.instance_variable_set(:@hooked, true)
  end
end
class Target
  include Hookable
end
puts Target.instance_variable_get(:@hooked)
