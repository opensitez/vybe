# vybe-test: ruby/functions/def_multi_params
# origin: languages/ruby/tests/ruby/test_functions.rs
# vybe-test-mode: compile

def calc(a, b, c = 0)
  a + b + c
end
