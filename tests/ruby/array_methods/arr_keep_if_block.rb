# vybe-test: ruby/array_methods/arr_keep_if_block
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

a = [1, 2, 3, 4]
a.keep_if { |n| n.even? }
