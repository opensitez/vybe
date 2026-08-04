# vybe-test: ruby/hash_methods/hash_dig
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = {'a' => {'b' => 42}}
x = h.dig('a', 'b')
