# vybe-test: ruby/enumerable/each_with_object_hash
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

x = ['a', 'b', 'c'].each_with_object({}) { |s, h| h[s] = s.upcase }
