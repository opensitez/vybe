# vybe-test: ruby/struct/struct_to_s_override
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Product = Struct.new(:name, :price) do
  def to_s
    name.to_s + ': $' + price.to_s
  end
end
puts Product.new("Widget", 9.99)
