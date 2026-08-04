# vybe-test: ruby/enumerable/all_with_block
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [2, 4, 6].all? { |n| n.even? }
