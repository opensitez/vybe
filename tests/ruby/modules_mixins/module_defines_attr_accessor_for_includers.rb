# vybe-test: ruby/modules_mixins/module_defines_attr_accessor_for_includers
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module HasName
  attr_accessor :name
end
class Product
  include HasName
end
p = Product.new
p.name = "Widget"
puts p.name
