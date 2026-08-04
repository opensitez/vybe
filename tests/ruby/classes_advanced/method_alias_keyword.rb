# vybe-test: ruby/classes_advanced/method_alias_keyword
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Talker
  def speak
    'speaking'
  end
  alias say speak
end
t = Talker.new
t.say
