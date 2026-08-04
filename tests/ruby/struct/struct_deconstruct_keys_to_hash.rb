# vybe-test: ruby/struct/struct_deconstruct_keys_to_hash
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Point3D = Struct.new(:x, :y, :z)
pt = Point3D.new(1, 2, 3)
h = pt.deconstruct_keys([:x, :z])
puts h.inspect
