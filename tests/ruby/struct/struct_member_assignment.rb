# vybe-test: ruby/struct/struct_member_assignment
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Box = Struct.new(:width, :height)
b = Box.new(10, 20)
b.width = 30
puts b.width
