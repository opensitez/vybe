# vybe-test: ruby/pattern_matching/case_in_string_match
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


val = "hello"
case val
in String => s
  puts "string: #{s}"
end
