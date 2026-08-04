# vybe-test: ruby/array_methods/arr_find_index_of_match
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = [10, 20, 30].find_index { |n| n == 20 }
