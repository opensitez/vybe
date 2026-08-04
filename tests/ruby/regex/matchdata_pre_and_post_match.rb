# vybe-test: ruby/regex/matchdata_pre_and_post_match
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

md = 'hello world'.match(/wor/)
pre = md.pre_match
post = md.post_match
