# vybe-test: ruby/array_methods/arr_min_by_block_value
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = ['banana', 'apple', 'cherry'].min_by { |s| s.length }
