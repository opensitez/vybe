# vybe-test: ruby/functions/recursion
# origin: languages/ruby/tests/ruby/test_functions.rs
# vybe-test-mode: compile

def fact(n)
  if n <= 1
    return 1
  end
  n * fact(n - 1)
end
