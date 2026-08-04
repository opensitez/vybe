# vybe-test: ruby/array_methods/arr_new_with_size_and_block
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = Array.new(5) { |i| i * 2 }
