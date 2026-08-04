# vybe-test: ruby/debug/minimal_break_in_while
# origin: languages/ruby/tests/ruby/test_debug.rs
# vybe-test-mode: compile

while true
break if true
end
