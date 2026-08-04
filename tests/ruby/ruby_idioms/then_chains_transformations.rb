# vybe-test: ruby/ruby_idioms/then_chains_transformations
# origin: languages/ruby/tests/ruby/test_ruby_idioms.rs
# vybe-test-mode: compile

result = 5.then { |x| x + 1 }.then { |x| x * 2 }
