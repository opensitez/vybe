# vybe-test: ruby/array_methods/arr_filter_map_combined
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = [1, 2, 3, 4, 5].filter_map { |n| n * 2 if n.odd? }
