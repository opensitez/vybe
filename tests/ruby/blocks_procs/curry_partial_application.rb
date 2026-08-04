# vybe-test: ruby/blocks_procs/curry_partial_application
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

add = ->(a, b) { a + b }
curried = add.curry
