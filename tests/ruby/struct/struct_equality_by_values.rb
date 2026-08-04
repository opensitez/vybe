# vybe-test: ruby/struct/struct_equality_by_values
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Vec = Struct.new(:x, :y)
a = Vec.new(1, 2)
b = Vec.new(1, 2)
c = Vec.new(3, 4)
puts a == b
puts a == c
