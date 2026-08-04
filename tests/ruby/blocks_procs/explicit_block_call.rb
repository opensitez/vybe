# vybe-test: ruby/blocks_procs/explicit_block_call
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

def run(&block)
  block.call(42)
end
run { |x| puts x }
