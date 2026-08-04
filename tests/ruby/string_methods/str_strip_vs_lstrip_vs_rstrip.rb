# vybe-test: ruby/string_methods/str_strip_vs_lstrip_vs_rstrip
# origin: languages/ruby/tests/ruby/test_string_methods.rs
# vybe-test-mode: compile


s = "  hello  "
a = s.strip
b = s.lstrip
c = s.rstrip
