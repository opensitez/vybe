# vybe-test: ruby/pattern_matching/pattern_nested_hash_and_array
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


response = { status: 200, body: ["ok", "done"] }
case response
in { status: 200, body: [first, *] }
  puts first
end
