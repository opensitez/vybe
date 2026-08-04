# vybe-test: ruby/pattern_matching/array_pattern_first_rest
# origin: languages/ruby/tests/ruby/test_pattern_matching.rs
# vybe-test-mode: compile


case [1, 2, 3]
in [first, *rest]
  puts first
  puts rest.inspect
end
