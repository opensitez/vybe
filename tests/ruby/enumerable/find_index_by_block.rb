# vybe-test: ruby/enumerable/find_index_by_block
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [10, 20, 30].find_index { |n| n > 15 }
