# vybe-test: ruby/classes_advanced/private_method_definition
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Vault
  def open
    unlock
  end
  private
  def unlock
    'unlocked'
  end
end
