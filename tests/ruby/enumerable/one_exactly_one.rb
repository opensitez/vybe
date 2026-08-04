# vybe-test: ruby/enumerable/one_exactly_one
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [1, 2, 3].one? { |n| n == 2 }
