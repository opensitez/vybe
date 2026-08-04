# vybe-test: ruby/regex/regexp_character_class
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

result = 'Hello World'.gsub(/[a-z]/, '*')
