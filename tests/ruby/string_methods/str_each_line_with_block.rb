# vybe-test: ruby/string_methods/str_each_line_with_block
# origin: languages/ruby/tests/ruby/test_string_methods.rs
# vybe-test-mode: compile

"line1\nline2\n".each_line { |l| puts l.chomp }
