# vybe-test: ruby/struct/struct_to_a_ordered_values
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Color = Struct.new(:r, :g, :b)
c = Color.new(255, 128, 0)
puts c.to_a.inspect
