# vybe-test: ruby/pattern_matching/hash_pattern_partial_match
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


event = { type: :click, x: 100, y: 200 }
case event
in { type: :click, x: Integer => x }
  puts 'click at x=' + x.to_s
end
