# vybe-test: ruby/blocks_procs/block_closure_captures_outer_var
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

x = 10
f = proc { puts x }
