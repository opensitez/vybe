# vybe-test: ruby/classes_advanced/hash_override
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Key
  def initialize(val)
    @val = val
  end
  def hash
    @val.hash
  end
end
k = Key.new(42)
k.hash
