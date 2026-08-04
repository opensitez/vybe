# vybe-test: ruby/blocks_procs/each_with_object_accumulator
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

result = [1, 2, 3].each_with_object([]) { |x, acc| acc.push(x * 2) }
