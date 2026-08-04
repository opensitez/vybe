# vybe-test: ruby/array_methods/arr_drop_while_condition
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = [1, 2, 3, 4, 5].drop_while { |n| n < 3 }
