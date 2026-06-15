use super::helpers::{compile_ok, run_ruby};

// ── FizzBuzz ────────────────────────────────────────────────────────────────

#[test]
fn fizzbuzz_compile() {
    compile_ok(
        r#"
for i in 1..15
  if i % 15 == 0
    puts "FizzBuzz"
  elsif i % 3 == 0
    puts "Fizz"
  elsif i % 5 == 0
    puts "Buzz"
  else
    puts i
  end
end
"#,
    );
}

#[test]
fn fizzbuzz_runtime() {
    let out = run_ruby(
        r#"
for i in 1..15
  if i % 15 == 0
    puts "FizzBuzz"
  elsif i % 3 == 0
    puts "Fizz"
  elsif i % 5 == 0
    puts "Buzz"
  else
    puts i
  end
end
"#,
    );
    assert_eq!(
        out,
        vec![
            "1", "2", "Fizz", "4", "Buzz", "Fizz", "7", "8", "Fizz", "Buzz", "11", "Fizz", "13",
            "14", "FizzBuzz"
        ]
    );
}

// ── Fibonacci ───────────────────────────────────────────────────────────────

#[test]
fn fibonacci_compile() {
    compile_ok(
        r#"
def fib(n)
  if n <= 1
    return n
  end
  fib(n - 1) + fib(n - 2)
end
puts fib(10)
"#,
    );
}

#[test]
fn fibonacci_runtime() {
    let out = run_ruby(
        r#"
def fib(n)
  if n <= 1
    return n
  end
  fib(n - 1) + fib(n - 2)
end
puts fib(10)
"#,
    );
    assert_eq!(out, vec!["55"]);
}

// ── Factorial ───────────────────────────────────────────────────────────────

#[test]
fn factorial_runtime() {
    let out = run_ruby(
        r#"
def factorial(n)
  if n <= 1
    return 1
  end
  n * factorial(n - 1)
end
puts factorial(10)
"#,
    );
    assert_eq!(out, vec!["3628800"]);
}

// ── Class with methods ──────────────────────────────────────────────────────

#[test]
fn class_program_compile() {
    compile_ok(
        r#"
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
"#,
    );
}

#[test]
fn class_program_runtime() {
    let out = run_ruby(
        r#"
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
"#,
    );
    assert_eq!(out, vec!["8"]);
}

// ── Array processing ────────────────────────────────────────────────────────

#[test]
fn array_processing_compile() {
    compile_ok(
        r#"
numbers = [5, 3, 8, 1, 9, 2]
sorted = numbers.sort
puts sorted.first
puts sorted.last
"#,
    );
}

#[test]
fn array_processing_runtime() {
    let out = run_ruby(
        r#"
numbers = [5, 3, 8, 1, 9, 2]
sorted = numbers.sort
puts sorted.first
puts sorted.last
"#,
    );
    assert_eq!(out, vec!["1", "9"]);
}

// ── String manipulation ─────────────────────────────────────────────────────

#[test]
fn string_program_runtime() {
    let out = run_ruby(
        r#"
sentence = "hello world ruby"
words = sentence.split(" ")
puts words.join("-")
"#,
    );
    assert_eq!(out, vec!["hello-world-ruby"]);
}

// ── Nested classes with inheritance ─────────────────────────────────────────

#[test]
fn inheritance_program_runtime() {
    let out = run_ruby(
        r#"
class Shape
  def initialize(name)
    @name = name
  end
  def describe
    puts @name
  end
end
class Circle < Shape
  def initialize(radius)
    super("Circle")
    @radius = radius
  end
  def area
    3.14 * @radius * @radius
  end
end
c = Circle.new(5)
c.describe
puts c.area
"#,
    );
    assert_eq!(out, vec!["Circle", "78.5"]);
}
