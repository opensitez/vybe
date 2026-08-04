# vybe-test: ruby/ruby_idioms/dunder_method_current_name
# origin: languages/ruby/tests/ruby/test_ruby_idioms.rs
# vybe-test-mode: compile

def my_func
  puts __method__
end
