# vybe-test: ruby/regex/regexp_dollar_ampersand_matched_string
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

'hello world' =~ /w\w+/
matched = $&
