# vybe-test: ruby/enumerable/flat_map_nested_array
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [[1, 2], [3, 4]].flat_map { |a| a }
