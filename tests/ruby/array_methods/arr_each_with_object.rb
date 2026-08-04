# vybe-test: ruby/array_methods/arr_each_with_object
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

[1, 2, 3].each_with_object([]) { |x, acc| acc.push(x * 2) }
