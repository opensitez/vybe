# vybe-test: ruby/blocks_procs/proc_new_from_block
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

p = Proc.new { |x| x + 1 }
