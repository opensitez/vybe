# vybe-test: ruby/ranges/range_each_slice
# origin: languages/ruby/tests/ruby/test_ranges.rs
# vybe-test-mode: compile

(1..9).each_slice(3) { |s| puts s.length }
