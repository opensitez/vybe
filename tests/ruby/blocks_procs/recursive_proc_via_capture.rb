# vybe-test: ruby/blocks_procs/recursive_proc_via_capture
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

fib = nil
fib = ->(n) { n < 2 ? n : fib.(n - 1) + fib.(n - 2) }
