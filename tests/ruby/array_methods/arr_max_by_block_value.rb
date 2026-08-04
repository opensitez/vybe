# vybe-test: ruby/array_methods/arr_max_by_block_value
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = ['banana', 'ap', 'cherry'].max_by { |s| s.length }
