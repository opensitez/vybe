# vybe-test: ruby/enumerable/take_while_condition
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [1, 2, 3, 4, 5].take_while { |n| n < 4 }
