# vybe-test: ruby/blocks_procs/lazy_map_first_n
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

result = (1..Float::INFINITY).lazy.map { |x| x * 2 }.first(5)
