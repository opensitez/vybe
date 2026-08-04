# vybe-test: ruby/array_methods/arr_any_predicate
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = [1, 2, 3].any? { |n| n > 2 }
