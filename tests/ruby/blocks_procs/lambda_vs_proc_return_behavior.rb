# vybe-test: ruby/blocks_procs/lambda_vs_proc_return_behavior
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

def test_lambda
  f = lambda { return 1 }
  f.call
  2
end
