# vybe-test: ruby/hash_methods/hash_filter_map
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = {'a' => 1, 'b' => 2, 'c' => 3}
x = h.filter_map { |k, v| [k, v * 2] if v > 1 }
