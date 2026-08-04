# vybe-test: ruby/modules_mixins/method_resolution_order_left_to_right
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


module A
  def hello; "A"; end
end
module B
  def hello; "B"; end
end
class C
  include A
  include B
end
puts C.new.hello
