# vybe-test: ruby/blocks_procs/then_pipes_value_through_block
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

result = 5.then { |x| x * 2 }
