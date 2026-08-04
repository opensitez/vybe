# vybe-test: ruby/regex/regexp_match_vs_string_match
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

r = /\d+/
md1 = r.match('abc 42 def')
md2 = 'abc 42 def'.match(r)
