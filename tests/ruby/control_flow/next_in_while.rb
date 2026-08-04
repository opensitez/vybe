# vybe-test: ruby/control_flow/next_in_while
# origin: languages/ruby/tests/ruby/test_control_flow.rs
# vybe-test-mode: compile

x = 0
while x < 10
  x += 1
  next if x % 2 == 0
  puts x
end
