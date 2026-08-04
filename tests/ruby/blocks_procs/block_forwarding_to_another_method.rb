# vybe-test: ruby/blocks_procs/block_forwarding_to_another_method
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

def outer(&block)
  inner(&block)
end
def inner
  yield 1
end
