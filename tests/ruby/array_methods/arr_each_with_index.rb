# vybe-test: ruby/array_methods/arr_each_with_index
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

[10, 20, 30].each_with_index { |v, i| puts i }
