# vybe-test: ruby/ruby_idioms/equal_vs_double_equals_distinction
# origin: languages/ruby/tests/ruby/test_ruby_idioms.rs
# vybe-test-mode: compile

a = 'hello'
b = 'hello'
value_eq = a == b
identity_eq = a.equal?(b)
