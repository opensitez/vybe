# vybe-test: ruby/pattern_matching/pin_operator_matches_variable
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


expected = 42
case [1, 42, 3]
in [*, ^expected, *]
  puts "found expected"
end
