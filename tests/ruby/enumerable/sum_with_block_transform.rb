# vybe-test: ruby/enumerable/sum_with_block_transform
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [1, 2, 3].sum { |n| n * 2 }
