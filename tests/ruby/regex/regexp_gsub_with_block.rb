# vybe-test: ruby/regex/regexp_gsub_with_block
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

result = 'hello world'.gsub(/\w+/) { |w| w.upcase }
