# vybe-test: ruby/blocks_procs/enumerator_lazy_chained
# origin: languages/ruby/tests/ruby/test_blocks_procs.rs
# vybe-test-mode: compile

result = [1, 2, 3, 4, 5].lazy.select { |x| x.odd? }.map { |x| x * 10 }.first(2)
