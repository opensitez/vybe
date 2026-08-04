# vybe-test: ruby/enumerable/none_with_block
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [1, 3, 5].none? { |n| n.even? }
