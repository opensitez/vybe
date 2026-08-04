# vybe-test: ruby/regex/regexp_anchors
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

a = 'hello' =~ /\Ahello\Z/
b = 'hello' =~ /^hello$/
