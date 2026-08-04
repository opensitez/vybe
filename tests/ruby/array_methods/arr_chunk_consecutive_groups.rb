# vybe-test: ruby/array_methods/arr_chunk_consecutive_groups
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

[1, 1, 2, 2, 3].chunk { |n| n }.to_a
