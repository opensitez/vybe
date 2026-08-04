# vybe-test: ruby/hash_methods/hash_all_q
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = {'a' => 1, 'b' => 2}
x = h.all? { |k, v| v > 0 }
