# vybe-test: ruby/blocks_procs/explicit_block_parameter
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

def run(&block)
  block.call
end
