# vybe-test: ruby/enumerable/each_slice_chunks
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

[1, 2, 3, 4, 5].each_slice(2) { |s| puts s.length }
