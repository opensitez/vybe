# vybe-test: ruby/ruby_idioms/clone_preserves_frozen_state
# origin: languages/ruby/tests/ruby/test_ruby_idioms.rs
# vybe-test-mode: compile

orig = 'hello'.freeze
copy = orig.clone
