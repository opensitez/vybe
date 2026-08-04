# vybe-test: ruby/hash_methods/hash_new_default_proc
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = Hash.new { |hash, key| hash[key] = key.upcase }
h['hello']
