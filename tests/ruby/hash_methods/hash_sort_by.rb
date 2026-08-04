# vybe-test: ruby/hash_methods/hash_sort_by
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = {'b' => 2, 'a' => 1}
x = h.sort_by { |k, v| k }
