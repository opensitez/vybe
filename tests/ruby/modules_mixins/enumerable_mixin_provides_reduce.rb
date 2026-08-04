# vybe-test: ruby/modules_mixins/enumerable_mixin_provides_reduce
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


class NumberSet
  include Enumerable
  def initialize(*ns); @ns = ns; end
  def each(&b); @ns.each(&b); end
end
puts NumberSet.new(1, 2, 3, 4, 5).reduce(:+)
