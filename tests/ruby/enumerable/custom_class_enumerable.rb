# vybe-test: ruby/enumerable/custom_class_enumerable
# origin: languages/ruby/tests/ruby/test_enumerable.rs
# vybe-test-mode: compile

module Enumerable
end
class NumberBag
  include Enumerable
def initialize(arr)
    @data = arr
  end
def each(&block)
    @data.each(&block)
  end
end
bag = NumberBag.new([1, 2, 3])
bag.each { |n| puts n }
