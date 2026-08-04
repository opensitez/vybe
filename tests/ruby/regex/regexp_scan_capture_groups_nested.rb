# vybe-test: ruby/regex/regexp_scan_capture_groups_nested
# origin: languages/ruby/tests/ruby/test_regex.rs
# vybe-test-mode: compile

result = 'one 1, two 2'.scan(/(\w+) (\d)/)
