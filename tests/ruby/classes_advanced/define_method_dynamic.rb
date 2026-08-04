# vybe-test: ruby/classes_advanced/define_method_dynamic
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Greeter
  ['hello', 'goodbye'].each do |word|
    define_method(word) do
      puts word
    end
  end
end
