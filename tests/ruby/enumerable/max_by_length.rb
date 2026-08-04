# vybe-test: ruby/enumerable/max_by_length
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = ['banana', 'fig', 'apple'].max_by { |s| s.length }
