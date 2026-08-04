# vybe-test: ruby/hash_methods/hash_nested_access
# origin: languages/ruby/tests/ruby/test_hash_methods.rs
# vybe-test-mode: compile

h = {'outer' => {'inner' => 42}}
x = h['outer']['inner']
