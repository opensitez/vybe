# vybe-test: ruby/array_methods/arr_find_first_matching
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = [1, 2, 3, 4].find { |n| n > 2 }
