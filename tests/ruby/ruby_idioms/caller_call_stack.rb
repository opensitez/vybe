# vybe-test: ruby/ruby_idioms/caller_call_stack
# origin: languages/ruby/tests/ruby/test_ruby_idioms.rs
# vybe-test-mode: compile

def deep
  caller
end
deep
