# vybe-test: ruby/builtins/arr_reduce
# origin: languages/ruby/tests/ruby/test_builtins.rs
# vybe-test-mode: compile

x = [1, 2, 3].reduce(0) { |sum, x| sum + x }
