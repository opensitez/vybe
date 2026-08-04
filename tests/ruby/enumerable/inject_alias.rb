# vybe-test: ruby/enumerable/inject_alias
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [1, 2, 3].inject(0) { |acc, n| acc + n }
