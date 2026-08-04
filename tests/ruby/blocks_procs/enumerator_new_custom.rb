# vybe-test: ruby/blocks_procs/enumerator_new_custom
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

e = Enumerator.new { |y| y << 1; y << 2; y << 3 }
