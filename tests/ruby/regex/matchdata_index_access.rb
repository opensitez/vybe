# vybe-test: ruby/regex/matchdata_index_access
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

md = 'hello'.match(/(e)(l+)/)
full = md[0]
first = md[1]
