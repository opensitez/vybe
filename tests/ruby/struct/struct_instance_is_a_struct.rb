# vybe-test: ruby/struct/struct_instance_is_a_struct
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Flag = Struct.new(:code)
f = Flag.new("US")
puts f.is_a?(Struct)
