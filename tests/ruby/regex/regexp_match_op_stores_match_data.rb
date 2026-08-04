# vybe-test: ruby/regex/regexp_match_op_stores_match_data
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

'hello' =~ /ell/
m = $~
