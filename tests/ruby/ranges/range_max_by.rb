# vybe-test: ruby/ranges/range_max_by
# origin: languages/ruby/tests/ruby/test_ranges.rs
# vybe-test-mode: compile

x = (1..5).max_by { |n| -n }
