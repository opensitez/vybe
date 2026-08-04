# vybe-test: ruby/symbols/symbol_to_proc_in_select
# origin: languages/ruby/tests/ruby/test_symbols.rs
# vybe-test-mode: compile

result = [1, 2, 3, nil, false].select(&:itself)
