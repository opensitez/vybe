# vybe-test: ruby/blocks_procs/proc_call_three_syntaxes
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

f = ->(x) { x * 2 }
a = f.call(3)
b = f.(3)
c = f[3]
