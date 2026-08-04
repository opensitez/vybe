# vybe-test: ruby/enumerable/chunk_while_consecutive
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [1, 2, 3, 5, 6, 10].chunk_while { |a, b| b == a + 1 }.to_a
