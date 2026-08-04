# vybe-test: ruby/blocks_procs/method_ref_as_proc
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

def double(x)
  x * 2
end
result = [1, 2, 3].map(&method(:double))
