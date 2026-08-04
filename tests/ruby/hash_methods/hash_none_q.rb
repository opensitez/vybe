# vybe-test: ruby/hash_methods/hash_none_q
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = {'a' => 1, 'b' => 2}
x = h.none? { |k, v| v > 5 }
