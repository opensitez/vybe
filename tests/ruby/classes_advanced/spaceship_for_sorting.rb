# vybe-test: ruby/classes_advanced/spaceship_for_sorting
# origin: languages/ruby/tests/ruby/test_classes_advanced.rs
# vybe-test-mode: compile

class Weight
  def initialize(kg)
    @kg = kg
  end
  def <=>(other)
    @kg <=> other.instance_variable_get(:@kg)
  end
end
weights = [Weight.new(5), Weight.new(2), Weight.new(8)]
weights.sort { |a, b| a <=> b }
