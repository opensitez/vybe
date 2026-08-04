# vybe-test: ruby/symbols/symbol_identity_same_object_id
# origin: languages/ruby/tests/ruby/test_symbols.rs
# vybe-test-mode: compile

result = :hello.object_id == :hello.object_id
