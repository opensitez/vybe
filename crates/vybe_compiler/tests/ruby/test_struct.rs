use super::helpers::compile_ok;

// ── Basic Struct creation ─────────────────────────────────────
#[test] fn struct_new_with_members() { compile_ok(r#"
Point = Struct.new(:x, :y)
p = Point.new(1, 2)
puts p.x
puts p.y
"#); }

#[test] fn struct_member_assignment() { compile_ok(r#"
Box = Struct.new(:width, :height)
b = Box.new(10, 20)
b.width = 30
puts b.width
"#); }

// ── Struct with keyword_init ──────────────────────────────────
#[test] fn struct_keyword_init() { compile_ok(r#"
Config = Struct.new(:host, :port, keyword_init: true)
c = Config.new(host: "localhost", port: 8080)
puts c.host
puts c.port
"#); }

// ── Struct.members ────────────────────────────────────────────
#[test] fn struct_members_returns_array() { compile_ok(r#"
Person = Struct.new(:name, :age, :email)
puts Person.members.inspect
"#); }

// ── Struct to_a and to_h ──────────────────────────────────────
#[test] fn struct_to_a_ordered_values() { compile_ok(r#"
Color = Struct.new(:r, :g, :b)
c = Color.new(255, 128, 0)
puts c.to_a.inspect
"#); }

#[test] fn struct_to_h_as_symbol_hash() { compile_ok(r#"
User = Struct.new(:name, :role)
u = User.new("Alice", :admin)
puts u.to_h.inspect
"#); }

// ── Struct equality ───────────────────────────────────────────
#[test] fn struct_equality_by_values() { compile_ok(r#"
Vec = Struct.new(:x, :y)
a = Vec.new(1, 2)
b = Vec.new(1, 2)
c = Vec.new(3, 4)
puts a == b
puts a == c
"#); }

// ── Struct with custom methods ────────────────────────────────
#[test] fn struct_block_adds_methods() { compile_ok(r#"
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
"#); }

#[test] fn struct_to_s_override() { compile_ok(r#"
Product = Struct.new(:name, :price) do
  def to_s
    name.to_s + ': $' + price.to_s
  end
end
puts Product.new("Widget", 9.99)
"#); }

// ── Struct as value object ────────────────────────────────────
#[test] fn struct_as_immutable_value_object() { compile_ok(r#"
Coordinate = Struct.new(:lat, :lon) do
  def to_s
    "(#{lat}, #{lon})"
  end
end
home = Coordinate.new(40.7128, -74.0060)
puts home
"#); }

// ── Struct in array sorting ───────────────────────────────────
#[test] fn struct_array_sort_by_member() { compile_ok(r#"
Employee = Struct.new(:name, :salary)
employees = [
  Employee.new("Bob", 50000),
  Employee.new("Alice", 75000),
  Employee.new("Carol", 60000)
]
sorted = employees.sort_by(&:salary)
puts sorted.map(&:name).inspect
"#); }

// ── Struct instance checks ────────────────────────────────────
#[test] fn struct_instance_is_a_struct() { compile_ok(r#"
Flag = Struct.new(:code)
f = Flag.new("US")
puts f.is_a?(Struct)
"#); }

// ── Struct with protected/private methods ────────────────────
#[test] fn struct_with_private_helper() { compile_ok(r#"
Rectangle = Struct.new(:w, :h) do
  def area; compute; end
  private
  def compute; w * h; end
end
puts Rectangle.new(4, 5).area
"#); }

// ── Struct enumerable ─────────────────────────────────────────
#[test] fn struct_each_iterates_values() { compile_ok(r#"
RGB = Struct.new(:r, :g, :b)
color = RGB.new(100, 150, 200)
color.each { |v| puts v }
"#); }

#[test] fn struct_map_transforms_values() { compile_ok(r#"
Pair = Struct.new(:a, :b)
p = Pair.new(3, 4)
doubled = p.map { |v| v * 2 }
puts doubled.inspect
"#); }

// ── Struct deconstruct (Ruby 3) ───────────────────────────────
#[test] fn struct_deconstruct_to_array() { compile_ok(r#"
Point3D = Struct.new(:x, :y, :z)
pt = Point3D.new(1, 2, 3)
arr = pt.deconstruct
puts arr.inspect
"#); }

#[test] fn struct_deconstruct_keys_to_hash() { compile_ok(r#"
Point3D = Struct.new(:x, :y, :z)
pt = Point3D.new(1, 2, 3)
h = pt.deconstruct_keys([:x, :z])
puts h.inspect
"#); }

// ── Struct inheritance ────────────────────────────────────────
#[test] fn struct_subclass_adds_methods() { compile_ok(r#"
Animal = Struct.new(:name, :sound)
class Dog < Animal
  def speak; name.to_s + ' says ' + sound.to_s + '!'; end
end
puts Dog.new("Rex", "woof").speak
"#); }

// ── Struct with respond_to? ───────────────────────────────────
#[test] fn struct_responds_to_member_methods() { compile_ok(r#"
Item = Struct.new(:id, :label)
item = Item.new(1, "thing")
puts item.respond_to?(:id)
puts item.respond_to?(:label)
puts item.respond_to?(:nonexistent)
"#); }

// ── Struct frozen instances ───────────────────────────────────
#[test] fn struct_can_be_frozen() { compile_ok(r#"
Token = Struct.new(:value)
t = Token.new("abc").freeze
puts t.frozen?
"#); }
