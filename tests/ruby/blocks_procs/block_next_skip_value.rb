# vybe-test: ruby/blocks_procs/block_next_skip_value
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

result = [1, 2, 3, 4].map { |x| next 0 if x.even?; x }
