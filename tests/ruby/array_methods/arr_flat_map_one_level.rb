# vybe-test: ruby/array_methods/arr_flat_map_one_level
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = [[1, 2], [3, 4]].flat_map { |a| a }
