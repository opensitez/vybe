# vybe-test: ruby/regex/regexp_sub_backreference_in_replacement
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

result = 'hello world'.sub(/(\w+)/, '[\1]')
