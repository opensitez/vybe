# vybe-test: ruby/array_methods/arr_detect_alias
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = [1, 2, 3].detect { |n| n.even? }
