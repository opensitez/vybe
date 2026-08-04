# vybe-test: ruby/classes_advanced/eql_override
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Vector
  def initialize(x, y)
    @x = x
    @y = y
  end
  def eql?(other)
    @x == other.instance_variable_get(:@x) && @y == other.instance_variable_get(:@y)
  end
end
v1 = Vector.new(1, 2)
v2 = Vector.new(1, 2)
v1.eql?(v2)
