# vybe-test: ruby/blocks_procs/freeze_object_then_use_in_block
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

s = 'hello'.freeze
result = [s].map { |x| x.length }
