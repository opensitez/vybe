# vybe-test: ruby/blocks_procs/block_closure_modifies_outer_var
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

total = 0
[1, 2, 3].each { |n| total += n }
