# vybe-test: ruby/regex/regexp_gsub_with_hash
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

result = 'aeiou'.gsub(/[aeiou]/, 'a' => '1', 'e' => '2', 'i' => '3')
