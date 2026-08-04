# vybe-test: ruby/ranges/range_any_q
# origin: languages/ruby/tests/ruby/test_ranges.rs
# vybe-test-mode: compile

x = (1..10).any? { |n| n > 8 }
