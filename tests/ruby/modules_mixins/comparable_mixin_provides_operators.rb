# vybe-test: ruby/modules_mixins/comparable_mixin_provides_operators
# origin: languages/ruby/tests/ruby/test_modules_mixins.rs
# vybe-test-mode: compile


class Box
  include Comparable
  attr_reader :volume
  def initialize(v); @volume = v; end
  def <=>(other); @volume <=> other.volume; end
end
small = Box.new(10)
large = Box.new(50)
puts small < large
puts large > small
puts small <= Box.new(10)
