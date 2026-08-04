# vybe-test: ruby/pattern_matching/array_pattern_fixed_length
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


case [10, 20]
in [a, b]
  puts a + b
end
