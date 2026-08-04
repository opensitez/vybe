# vybe-test: ruby/enumerable/filter_map_combined
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [1, 2, 3, 4, 5].filter_map { |n| n * 2 if n.odd? }
