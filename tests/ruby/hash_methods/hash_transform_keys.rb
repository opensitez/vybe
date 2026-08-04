# vybe-test: ruby/hash_methods/hash_transform_keys
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = {'a' => 1, 'b' => 2}
x = h.transform_keys { |k| k.upcase }
