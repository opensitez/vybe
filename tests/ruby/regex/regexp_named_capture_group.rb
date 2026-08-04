# vybe-test: ruby/regex/regexp_named_capture_group
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

m = 'John 30'.match(/(?<name>\w+) (?<age>\d+)/)
