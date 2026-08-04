# vybe-test: ruby/regex/regexp_capture_group_dollar_vars
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

'2024-01-15' =~ /(\d{4})-(\d{2})/
y = $1
m = $2
