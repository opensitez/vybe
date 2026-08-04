# vybe-test: ruby/string_methods/str_gsub_with_regex_pattern
# origin: languages/ruby/tests/ruby/test_string_methods.rs
# vybe-test-mode: compile

x = 'hello world'.gsub(/[aeiou]/, '*')
