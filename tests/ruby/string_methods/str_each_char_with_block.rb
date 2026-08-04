# vybe-test: ruby/string_methods/str_each_char_with_block
# origin: languages/ruby/tests/ruby/test_string_methods.rs
# vybe-test-mode: compile

'abc'.each_char { |c| puts c }
