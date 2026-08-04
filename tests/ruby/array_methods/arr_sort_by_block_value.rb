# vybe-test: ruby/array_methods/arr_sort_by_block_value
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = ['banana', 'apple', 'cherry'].sort_by { |s| s.length }
