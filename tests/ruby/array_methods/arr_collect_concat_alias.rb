# vybe-test: ruby/array_methods/arr_collect_concat_alias
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

x = [[1, 2], [3]].collect_concat { |a| a }
