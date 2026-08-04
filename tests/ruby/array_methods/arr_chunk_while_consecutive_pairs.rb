# vybe-test: ruby/array_methods/arr_chunk_while_consecutive_pairs
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

[1, 2, 3, 5, 6, 10].chunk_while { |a, b| b == a + 1 }.to_a
