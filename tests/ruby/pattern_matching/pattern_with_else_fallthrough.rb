# vybe-test: ruby/pattern_matching/pattern_with_else_fallthrough
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


case { type: :unknown }
in { type: :click }
  puts "click"
in { type: :keypress }
  puts "keypress"
else
  puts "other"
end
