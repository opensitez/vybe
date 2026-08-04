# vybe-test: ruby/struct/struct_block_adds_methods
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Circle = Struct.new(:radius) do
  def area
    Math::PI * radius ** 2
  end
  def circumference
    2 * Math::PI * radius
  end
end
c = Circle.new(5)
puts c.area.round(2)
