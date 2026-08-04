# vybe-test: ruby/control_flow/loop_stmt
# origin: languages/ruby/tests/ruby/test_control_flow.rs
# vybe-test-mode: compile

i = 0
loop do
  break if i >= 3
  i += 1
end
