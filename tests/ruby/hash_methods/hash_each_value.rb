# vybe-test: ruby/hash_methods/hash_each_value
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = {'a' => 1, 'b' => 2}
h.each_value { |v| puts v }
