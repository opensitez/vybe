# vybe-test: ruby/pattern_matching/custom_deconstruct_for_array_pattern
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


class Point
  attr_reader :x, :y
  def initialize(x, y); @x = x; @y = y; end
  def deconstruct; [@x, @y]; end
end
case Point.new(3, 4)
in [x, y]
  puts x + y
end
