# vybe-test: ruby/hash_methods/hash_each_pair
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = {'x' => 10}
h.each_pair { |k, v| puts v }
