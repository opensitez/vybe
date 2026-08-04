# vybe-test: ruby/ranges/range_step_float
# origin: languages/ruby/tests/ruby/test_ranges.rs
# vybe-test-mode: compile

(0.0..1.0).step(0.25) { |f| puts f }
