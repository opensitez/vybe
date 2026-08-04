# vybe-test: ruby/struct/struct_responds_to_member_methods
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Item = Struct.new(:id, :label)
item = Item.new(1, "thing")
puts item.respond_to?(:id)
puts item.respond_to?(:label)
puts item.respond_to?(:nonexistent)
