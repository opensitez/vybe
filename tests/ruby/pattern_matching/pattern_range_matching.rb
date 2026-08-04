# vybe-test: ruby/pattern_matching/pattern_range_matching
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


score = 85
case score
in 90..100
  puts "A"
in 80...90
  puts "B"
in 70...80
  puts "C"
else
  puts "F"
end
