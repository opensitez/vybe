# vybe-test: ruby/struct/struct_to_h_as_symbol_hash
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


User = Struct.new(:name, :role)
u = User.new("Alice", :admin)
puts u.to_h.inspect
