# vybe-test: ruby/array_methods/arr_each_slice_chunks
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

[1, 2, 3, 4, 5, 6].each_slice(2) { |s| puts s.length }
