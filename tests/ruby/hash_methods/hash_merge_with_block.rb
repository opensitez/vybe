# vybe-test: ruby/hash_methods/hash_merge_with_block
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h1 = {'a' => 1}
h2 = {'a' => 2}
x = h1.merge(h2) { |key, old, new_v| old + new_v }
