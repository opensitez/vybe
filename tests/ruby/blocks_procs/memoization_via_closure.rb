# vybe-test: ruby/blocks_procs/memoization_via_closure
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

cache = {}
memo = ->(n) { cache[n] ||= n * n }
