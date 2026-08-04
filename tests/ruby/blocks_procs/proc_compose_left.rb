# vybe-test: ruby/blocks_procs/proc_compose_left
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

double = ->(x) { x * 2 }
increment = ->(x) { x + 1 }
inc_then_double = double << increment
