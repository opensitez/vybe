# vybe-test: ruby/builtins/arr_map
# origin: languages/ruby/tests/ruby/test_builtins.rs
# vybe-test-mode: compile

x = [1, 2, 3].map { |x| x * 2 }
