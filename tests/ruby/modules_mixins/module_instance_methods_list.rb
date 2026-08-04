# vybe-test: ruby/modules_mixins/module_instance_methods_list
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module Tools
  def hammer; "bang"; end
  def screwdriver; "turn"; end
end
puts Tools.instance_methods.include?(:hammer)
