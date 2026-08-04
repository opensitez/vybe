# vybe-test: ruby/enumerable/min_by_length
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = ['banana', 'fig', 'apple'].min_by { |s| s.length }
