# vybe-test: ruby/blocks_procs/proc_arity_method
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

f = ->(a, b) { a + b }
n = f.arity
