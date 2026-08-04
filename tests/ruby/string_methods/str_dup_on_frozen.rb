# vybe-test: ruby/string_methods/str_dup_on_frozen
# origin: languages/ruby/tests/ruby/test_string_methods.rs
# vybe-test-mode: compile

x = 'hello'.freeze
y = x.dup
