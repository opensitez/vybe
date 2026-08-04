# vybe-test: ruby/modules_mixins/forwardable_delegates_methods
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


require 'forwardable'
class Stack
  extend Forwardable
  def_delegators :@data, :push, :pop, :size, :empty?
  def initialize; @data = []; end
end
s = Stack.new
s.push(1)
s.push(2)
puts s.size
puts s.pop
