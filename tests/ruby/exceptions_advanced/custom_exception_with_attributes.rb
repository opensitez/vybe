# vybe-test: ruby/exceptions_advanced/custom_exception_with_attributes
# origin: languages/ruby/tests/ruby/test_exceptions_advanced.rs
# vybe-test-mode: compile

class HttpError < StandardError
  attr_reader :code
  def initialize(msg, code)
    super(msg)
    @code = code
  end
end
begin
  raise HttpError.new('not found', 404)
rescue HttpError => e
  puts e.code
end
