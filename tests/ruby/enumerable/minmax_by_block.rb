# vybe-test: ruby/enumerable/minmax_by_block
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = ['banana', 'fig', 'apple'].minmax_by { |s| s.length }
