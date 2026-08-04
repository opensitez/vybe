# vybe-test: ruby/symbols/symbol_hash_key_vs_string_key
# origin: languages/ruby/tests/ruby/test_symbols.rs
# vybe-test-mode: compile

h = {}
h[:foo] = 1
h['foo'] = 2
result = h[:foo] == h['foo']
