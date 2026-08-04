# vybe-test: ruby/blocks_procs/curry_apply_remaining_args
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

add = ->(a, b) { a + b }
add5 = add.curry.(5)
result = add5.(3)
