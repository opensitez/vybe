# vybe-test: ruby/symbols/symbol_to_proc_in_reject
# origin: languages/ruby/tests/ruby/test_symbols.rs
# vybe-test-mode: compile

result = ['a', '', 'b', ''].reject(&:empty?)
