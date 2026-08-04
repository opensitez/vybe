# vybe-test: ruby/enumerable/each_cons_window
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

[1, 2, 3, 4, 5].each_cons(3) { |c| puts c.length }
