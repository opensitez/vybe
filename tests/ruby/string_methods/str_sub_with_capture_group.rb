# vybe-test: ruby/string_methods/str_sub_with_capture_group
# origin: languages/ruby/tests/ruby/test_string_methods.rs
# vybe-test-mode: compile

x = 'hello world'.sub(/(\w+)/, 'HI')
