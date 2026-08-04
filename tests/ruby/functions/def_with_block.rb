# vybe-test: ruby/functions/def_with_block
# origin: languages/ruby/tests/ruby/test_functions.rs
# vybe-test-mode: compile

def each_item(&block)
  block.call(1)
end
