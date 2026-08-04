# vybe-test: ruby/pattern_matching/find_pattern_in_array
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


case [1, 2, 42, 3, 4]
in [*, 42, *]
  puts "found 42"
end
