# vybe-test: ruby/builtins/hash_delete
# origin: languages/ruby/tests/ruby/test_builtins.rs
# vybe-test-mode: compile

h = { 'a' => 1 }
h.delete('a')
