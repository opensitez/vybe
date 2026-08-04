# vybe-test: ruby/enumerable/any_with_block
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [1, 2, 3].any? { |n| n > 2 }
