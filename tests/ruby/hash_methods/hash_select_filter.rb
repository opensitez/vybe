# vybe-test: ruby/hash_methods/hash_select_filter
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = {'a' => 1, 'b' => 2, 'c' => 3}
x = h.select { |k, v| v > 1 }
