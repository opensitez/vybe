# vybe-test: ruby/blocks_procs/block_given_check
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

def maybe_yield
  if block_given?
    yield
  else
    puts 'no block'
  end
end
