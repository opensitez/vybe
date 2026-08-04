# vybe-test: ruby/debug/full_while_break
# origin: languages/ruby/tests/ruby/test_debug.rs
# vybe-test-mode: compile

x = 0
while true
  break if x >= 3
  x += 1
end
