# vybe-test: ruby/classes_advanced/inspect_override_in_class
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Node
  def initialize(val)
    @val = val
  end
  def inspect
    'Node(' + @val.to_s + ')'
  end
end
n = Node.new(7)
n.inspect
