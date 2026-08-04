# vybe-test: ruby/programs/class_program_compile
# origin: languages/ruby/tests/ruby/test_programs.rs
# vybe-test-mode: compile


class Calculator
  def initialize(value)
    @value = value
  end

  def add(n)
    @value = @value + n
  end

  def result
    @value
  end
end

c = Calculator.new(0)
c.add(5)
c.add(3)
puts c.result
