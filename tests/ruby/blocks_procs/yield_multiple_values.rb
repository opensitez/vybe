# vybe-test: ruby/blocks_procs/yield_multiple_values
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

def pair
  yield 'key', 'value'
end
