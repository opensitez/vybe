# vybe-test: ruby/hash_methods/hash_new_default_value
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = Hash.new(0)
h['missing']
