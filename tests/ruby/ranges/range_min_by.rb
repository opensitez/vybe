# vybe-test: ruby/ranges/range_min_by
# origin: languages/ruby/tests/ruby/test_ranges.rs
# vybe-test-mode: compile

x = (1..5).min_by { |n| -n }
