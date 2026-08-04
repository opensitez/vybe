# vybe-test: ruby/symbols/symbol_as_send_method_name
# origin: languages/ruby/tests/ruby/test_symbols.rs
# vybe-test-mode: compile

result = 'hello'.send(:upcase)
