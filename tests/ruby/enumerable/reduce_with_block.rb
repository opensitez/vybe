# vybe-test: ruby/enumerable/reduce_with_block
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [1, 2, 3, 4].reduce { |sum, n| sum + n }
