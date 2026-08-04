# vybe-test: ruby/struct/struct_with_private_helper
# origin: languages/ruby/tests/ruby/test_struct.rs
# vybe-test-mode: compile


Rectangle = Struct.new(:w, :h) do
  def area; compute; end
  private
  def compute; w * h; end
end
puts Rectangle.new(4, 5).area
