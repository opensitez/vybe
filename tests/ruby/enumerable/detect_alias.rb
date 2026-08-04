# vybe-test: ruby/enumerable/detect_alias
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = [1, 2, 3, 4].detect { |n| n.even? }
