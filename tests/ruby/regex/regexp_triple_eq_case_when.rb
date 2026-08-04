# vybe-test: ruby/regex/regexp_triple_eq_case_when
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

s = 'hello123'
result = case s
when /^\d/ then 'digit'
when /^[a-z]/ then 'letter'
else 'other'
end
