# vybe-test: ruby/array_methods/arr_all_predicate
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = [2, 4, 6].all? { |n| n.even? }
