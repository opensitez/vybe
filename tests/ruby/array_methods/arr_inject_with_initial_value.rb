# vybe-test: ruby/array_methods/arr_inject_with_initial_value
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = [1, 2, 3].inject(10) { |sum, n| sum + n }
