# vybe-test: ruby/regex/regexp_flag_case_insensitive
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

r = /hello/i
x = 'HELLO' =~ r
