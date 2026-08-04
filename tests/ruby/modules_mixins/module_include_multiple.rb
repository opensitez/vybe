# vybe-test: ruby/modules_mixins/module_include_multiple
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module Swimmable
  def swim; "splashing"; end
end
module Flyable
  def fly; "soaring"; end
end
class Duck
  include Swimmable
  include Flyable
end
d = Duck.new
puts d.swim
puts d.fly
