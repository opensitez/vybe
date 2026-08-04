# vybe-test: ruby/enumerable/sort_by_string_length
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = ['banana', 'apple', 'fig'].sort_by { |s| s.length }
