# vybe-test: ruby/builtins/arr_reject
# origin: languages/ruby/tests/ruby/test_builtins.rs
# vybe-test-mode: compile

x = [1, 2, 3].reject { |x| x == 2 }
