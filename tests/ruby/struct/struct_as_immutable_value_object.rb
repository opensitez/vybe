# vybe-test: ruby/struct/struct_as_immutable_value_object
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Coordinate = Struct.new(:lat, :lon) do
  def to_s
    "(#{lat}, #{lon})"
  end
end
home = Coordinate.new(40.7128, -74.0060)
puts home
