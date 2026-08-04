# vybe-test: ruby/array_methods/arr_minmax_by_block
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = ['apple', 'fig', 'cherry'].minmax_by { |s| s.length }
