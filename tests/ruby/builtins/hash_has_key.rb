# vybe-test: ruby/builtins/hash_has_key
# origin: languages/ruby/tests/ruby/test_builtins.rs
# vybe-test-mode: compile

h = { 'a' => 1 }
x = h.has_key?('a')
