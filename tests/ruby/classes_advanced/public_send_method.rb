# vybe-test: ruby/classes_advanced/public_send_method
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Foo
  def greet
    'hello'
  end
end
f = Foo.new
f.public_send(:greet)
