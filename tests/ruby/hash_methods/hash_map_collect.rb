# vybe-test: ruby/hash_methods/hash_map_collect
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = {'a' => 1, 'b' => 2}
x = h.map { |k, v| [k, v * 2] }
