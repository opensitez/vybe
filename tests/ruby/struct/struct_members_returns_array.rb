# vybe-test: ruby/struct/struct_members_returns_array
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Person = Struct.new(:name, :age, :email)
puts Person.members.inspect
