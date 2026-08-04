# vybe-test: ruby/blocks_procs/respond_to_method_check
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

x = 'hello'
puts x.respond_to?(:upcase)
