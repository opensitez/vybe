# vybe-test: ruby/struct/struct_each_iterates_values
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


RGB = Struct.new(:r, :g, :b)
color = RGB.new(100, 150, 200)
color.each { |v| puts v }
