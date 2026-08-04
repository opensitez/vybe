# vybe-test: ruby/exceptions_advanced/throw_catch_flow_control
# origin: languages/ruby/tests/ruby/test_exceptions_advanced.rs
# vybe-test-mode: compile

result = catch(:done) do
  [1, 2, 3].each do |n|
    throw :done, n if n == 2
  end
end
