# vybe-test: ruby/ruby_idioms/tap_returns_receiver
# origin: languages/ruby/tests/ruby/test_ruby_idioms.rs
# vybe-test-mode: compile

[1, 2, 3].tap { |a| puts a.length }
