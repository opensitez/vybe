# vybe-test: ruby/modules_mixins/module_include_chain_calls_super
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module Base
  def info; "base"; end
end
module Derived
  include Base
  def info; super + "+derived"; end
end
class Thing
  include Derived
end
puts Thing.new.info
