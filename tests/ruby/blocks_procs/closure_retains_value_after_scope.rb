# vybe-test: ruby/blocks_procs/closure_retains_value_after_scope
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

def make_adder(n)
  ->(x) { x + n }
end
add10 = make_adder(10)
