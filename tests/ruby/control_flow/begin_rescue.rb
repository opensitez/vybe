# vybe-test: ruby/control_flow/begin_rescue
# origin: languages/ruby/tests/ruby/test_control_flow.rs
# vybe-test-mode: compile

begin
  x = 1 / 0
rescue
  puts 'error'
end
