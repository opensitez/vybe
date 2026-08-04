# vybe-test: ruby/exceptions_advanced/custom_exception_custom_message
# origin: languages/ruby/tests/ruby/test_exceptions_advanced.rs
# vybe-test-mode: compile

class AppError < StandardError
  def initialize(msg = 'app error occurred')
    super(msg)
  end
end
raise AppError rescue nil
