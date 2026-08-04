# vybe-test: ruby/ruby_idioms/safe_navigation_on_nil_returns_nil
# origin: languages/ruby/tests/ruby/test_ruby_idioms.rs
# vybe-test-mode: compile

s = nil
result = s&.upcase
