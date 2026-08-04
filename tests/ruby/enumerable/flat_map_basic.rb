# vybe-test: ruby/enumerable/flat_map_basic
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [1, 2, 3].flat_map { |n| [n, n * 2] }
