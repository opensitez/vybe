# vybe-test: ruby/ruby_idioms/dup_shallow_copy_not_frozen
# origin: languages/ruby/tests/ruby/test_ruby_idioms.rs
# vybe-test-mode: compile

orig = 'hello'.freeze
copy = orig.dup
puts copy.frozen?
