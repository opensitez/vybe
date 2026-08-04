# vybe-test: ruby/struct/struct_can_be_frozen
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Token = Struct.new(:value)
t = Token.new("abc").freeze
puts t.frozen?
