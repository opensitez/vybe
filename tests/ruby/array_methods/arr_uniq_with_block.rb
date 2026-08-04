# vybe-test: ruby/array_methods/arr_uniq_with_block
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = ['apple', 'Banana', 'cherry'].uniq { |s| s.downcase[0] }
