# vybe-test: ruby/ranges/range_select
# origin: languages/ruby/tests/ruby/test_ranges.rs
# vybe-test-mode: compile

x = (1..10).select { |i| i.even? }
