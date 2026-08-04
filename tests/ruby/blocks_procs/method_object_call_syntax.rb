# vybe-test: ruby/blocks_procs/method_object_call_syntax
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

def square(x)
  x * x
end
m = method(:square)
result = m.(5)
