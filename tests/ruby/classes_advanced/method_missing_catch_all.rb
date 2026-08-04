# vybe-test: ruby/classes_advanced/method_missing_catch_all
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Ghost
  def method_missing(name, *args)
    'called ' + name.to_s
  end
end
g = Ghost.new
g.anything
