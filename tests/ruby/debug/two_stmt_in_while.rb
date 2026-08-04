# vybe-test: ruby/debug/two_stmt_in_while
# origin: languages/ruby/tests/ruby/test_debug.rs
# vybe-test-mode: compile

while true
x = 1
break if true
end
