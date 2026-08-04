# vybe-test: ruby/struct/struct_array_sort_by_member
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Employee = Struct.new(:name, :salary)
employees = [
  Employee.new("Bob", 50000),
  Employee.new("Alice", 75000),
  Employee.new("Carol", 60000)
]
sorted = employees.sort_by(&:salary)
puts sorted.map(&:name).inspect
