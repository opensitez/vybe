# vybe-test: ruby/blocks_procs/tap_chain_side_effect
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

result = [1, 2, 3].tap { |a| puts a.length }.map { |x| x * 2 }
