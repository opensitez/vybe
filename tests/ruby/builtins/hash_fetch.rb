# vybe-test: ruby/builtins/hash_fetch
# origin: languages/ruby/tests/ruby/test_builtins.rs
# vybe-test-mode: compile

h = { 'a' => 1 }
x = h.fetch('a')
