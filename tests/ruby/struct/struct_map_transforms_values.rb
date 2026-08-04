# vybe-test: ruby/struct/struct_map_transforms_values
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Pair = Struct.new(:a, :b)
p = Pair.new(3, 4)
doubled = p.map { |v| v * 2 }
puts doubled.inspect
