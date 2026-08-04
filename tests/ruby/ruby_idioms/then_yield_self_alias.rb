# vybe-test: ruby/ruby_idioms/then_yield_self_alias
# origin: languages/ruby/tests/ruby/test_ruby_idioms.rs
# vybe-test-mode: compile

result = 5.yield_self { |x| x * 2 }
