# vybe-test: ruby/control_flow/begin_rescue_ensure
# origin: languages/ruby/tests/ruby/test_control_flow.rs
# vybe-test-mode: compile

begin
  x = 1
rescue => e
  puts e
ensure
  puts 'done'
end
