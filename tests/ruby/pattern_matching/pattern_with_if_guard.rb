# vybe-test: ruby/pattern_matching/pattern_with_if_guard
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


case 15
in n if n > 10
  puts "big: #{n}"
end
