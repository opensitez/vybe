# vybe-test: ruby/enumerable/group_by_even_odd
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [1, 2, 3, 4, 5, 6].group_by { |n| n % 2 == 0 ? 'even' : 'odd' }
