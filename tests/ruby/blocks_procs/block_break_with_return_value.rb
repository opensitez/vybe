# vybe-test: ruby/blocks_procs/block_break_with_return_value
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

result = [1, 2, 3, 4].each { |x| break x if x > 2 }
