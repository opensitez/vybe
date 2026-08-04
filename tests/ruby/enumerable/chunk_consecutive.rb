# vybe-test: ruby/enumerable/chunk_consecutive
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [1, 1, 2, 2, 3].chunk { |n| n }.to_a
