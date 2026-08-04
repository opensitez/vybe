# vybe-test: ruby/blocks_procs/block_yield_return_value_used
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

def transform(x)
  yield x
end
result = transform(5) { |n| n * 3 }
