# vybe-test: ruby/pattern_matching/hash_pattern_extracts_keys
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


data = { name: "Alice", age: 30 }
case data
in { name: String => name, age: Integer => age }
  puts name.to_s + ' is ' + age.to_s
end
