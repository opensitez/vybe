# vybe-test: ruby/struct/struct_new_with_members
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Point = Struct.new(:x, :y)
p = Point.new(1, 2)
puts p.x
puts p.y
