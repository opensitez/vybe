# vybe-test: ruby/blocks_procs/multiple_yield_calls
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

def three_times
  yield 1
  yield 2
  yield 3
end
