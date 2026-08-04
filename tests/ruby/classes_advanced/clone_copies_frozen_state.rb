# vybe-test: ruby/classes_advanced/clone_copies_frozen_state
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Tag
  attr_reader :name
  def initialize(n)
    @name = n
  end
end
t = Tag.new('x')
t.freeze
c = t.clone
c.frozen?
