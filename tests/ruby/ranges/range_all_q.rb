# vybe-test: ruby/ranges/range_all_q
# origin: languages/ruby/tests/ruby/test_ranges.rs
# vybe-test-mode: compile

x = (1..5).all? { |n| n > 0 }
