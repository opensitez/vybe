# vybe-test: ruby/hash_methods/hash_min_by
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = {'a' => 3, 'b' => 1, 'c' => 2}
x = h.min_by { |k, v| v }
