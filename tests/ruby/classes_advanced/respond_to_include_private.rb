# vybe-test: ruby/classes_advanced/respond_to_include_private
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Vault
  private
  def secret
    'hidden'
  end
end
v = Vault.new
v.respond_to?(:secret, true)
