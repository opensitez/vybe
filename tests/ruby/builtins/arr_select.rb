# vybe-test: ruby/builtins/arr_select
# origin: languages/ruby/tests/ruby/test_builtins.rs
# vybe-test-mode: compile

x = [1, 2, 3, 4].select { |x| x > 2 }
