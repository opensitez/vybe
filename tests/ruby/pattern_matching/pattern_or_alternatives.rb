# vybe-test: ruby/pattern_matching/pattern_or_alternatives
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


[1, 2, 3].each do |n|
  case n
  in 1 | 3
    puts "odd"
  in 2
    puts "even"
  end
end
