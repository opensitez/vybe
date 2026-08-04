# vybe-test: ruby/control_flow/break_in_while
# origin: languages/ruby/tests/ruby/test_control_flow.rs
# vybe-test-mode: compile

x = 0
while true
  break if x >= 3
  x += 1
end
