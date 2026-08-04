# vybe-test: ruby/control_flow/case_when
# origin: languages/ruby/tests/ruby/test_control_flow.rs
# vybe-test-mode: compile

x = 2
case x
when 1
  puts 'one'
when 2
  puts 'two'
else
  puts 'other'
end
