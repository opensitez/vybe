# vybe-test: ruby/symbols/symbol_to_proc_in_sort_by
# origin: languages/ruby/tests/ruby/test_symbols.rs
# vybe-test-mode: compile

words = ['banana', 'apple', 'cherry']
result = words.sort_by(&:length)
