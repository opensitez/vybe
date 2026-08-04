# vybe-test: ruby/regex/matchdata_named_access
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

md = 'John'.match(/(?<first>\w+)/)
n = md[:first]
