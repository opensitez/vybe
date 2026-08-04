# vybe-test: ruby/array_methods/arr_each_cons_sliding_window
# origin: languages/ruby/tests/ruby/test_array_methods.rs
# vybe-test-mode: compile

[1, 2, 3, 4, 5].each_cons(3) { |w| puts w.first }
