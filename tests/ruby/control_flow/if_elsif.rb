# vybe-test: ruby/control_flow/if_elsif
# origin: languages/ruby/tests/ruby/test_control_flow.rs
# vybe-test-mode: compile

x = 2
if x == 1
  puts 'a'
elsif x == 2
  puts 'b'
else
  puts 'c'
end
