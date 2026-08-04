# vybe-test: ruby/pattern_matching/array_pattern_nested
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


case [[1, 2], [3, 4]]
in [[a, b], [c, d]]
  puts a + b + c + d
end
