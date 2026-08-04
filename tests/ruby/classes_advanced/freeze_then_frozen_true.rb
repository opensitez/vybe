# vybe-test: ruby/classes_advanced/freeze_then_frozen_true
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Token
  attr_reader :val
  def initialize(v)
    @val = v
  end
end
t = Token.new('abc')
t.freeze
t.frozen?
