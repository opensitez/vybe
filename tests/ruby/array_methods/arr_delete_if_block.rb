# vybe-test: ruby/array_methods/arr_delete_if_block
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

a = [1, 2, 3, 4]
a.delete_if { |n| n.even? }
