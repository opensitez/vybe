# vybe-test: ruby/functions/def_defaults
# origin: languages/ruby/tests/ruby/test_functions.rs
# vybe-test-mode: compile

def greet(name = 'world')
  puts name
end
