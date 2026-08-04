# vybe-test: ruby/hash_methods/hash_compact
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = {'a' => 1, 'b' => nil, 'c' => 3}
x = h.compact
