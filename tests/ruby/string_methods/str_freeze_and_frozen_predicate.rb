# vybe-test: ruby/string_methods/str_freeze_and_frozen_predicate
# origin: languages/ruby/tests/ruby/test_string_methods.rs
# vybe-test-mode: compile

x = 'hello'.freeze
y = x.frozen?
