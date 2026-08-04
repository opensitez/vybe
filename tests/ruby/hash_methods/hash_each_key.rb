# vybe-test: ruby/hash_methods/hash_each_key
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = {'a' => 1, 'b' => 2}
h.each_key { |k| puts k }
