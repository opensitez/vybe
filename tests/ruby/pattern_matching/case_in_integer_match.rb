# vybe-test: ruby/pattern_matching/case_in_integer_match
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


value = 42
case value
in Integer => n
  puts "integer: #{n}"
end
