# vybe-test: ruby/ranges/range_none_q
# origin: languages/ruby/tests/ruby/test_ranges.rs
# vybe-test-mode: compile

x = (1..5).none? { |n| n > 10 }
