use vybe_parser_ruby::parse;
use vybe_compiler_ruby::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

// ── FizzBuzz ───────────────────────────────────────────────
#[test]
fn fizzbuzz() {
    compile_ok(r#"
for i in [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]
  if i % 15 == 0
    puts 'FizzBuzz'
  elsif i % 3 == 0
    puts 'Fizz'
  elsif i % 5 == 0
    puts 'Buzz'
  else
    puts i
  end
end
"#);
}

// ── Calculator class ───────────────────────────────────────
#[test]
fn calculator() {
    compile_ok(r#"
class Calculator
  def initialize
    @result = 0
  end

  def add(x)
    @result += x
  end

  def subtract(x)
    @result -= x
  end

  def result
    @result
  end
end

calc = Calculator.new
calc.add(10)
calc.subtract(3)
puts calc.result
"#);
}

// ── Linked list ────────────────────────────────────────────
#[test]
fn linked_list() {
    compile_ok(r#"
class Node
  attr_accessor :value, :next_node

  def initialize(value)
    @value = value
    @next_node = nil
  end
end

class LinkedList
  def initialize
    @head = nil
  end

  def push(value)
    node = Node.new(value)
    node.next_node = @head
    @head = node
  end
end

list = LinkedList.new
list.push(1)
list.push(2)
list.push(3)
"#);
}

// ── Stack class ────────────────────────────────────────────
#[test]
fn stack_impl() {
    compile_ok(r#"
class Stack
  def initialize
    @data = []
  end

  def push(item)
    @data.push(item)
  end

  def pop
    @data.pop
  end

  def empty?
    @data.empty?
  end
end

s = Stack.new
s.push(1)
s.push(2)
s.push(3)
s.pop
"#);
}

// ── Fibonacci with memoization ────────────────────────────
#[test]
fn memoized_fib() {
    compile_ok(r#"
def fib(n)
  return n if n <= 1
  fib(n - 1) + fib(n - 2)
end

result = fib(10)
puts result
"#);
}

// ── Sorting ────────────────────────────────────────────────
#[test]
fn sorting() {
    compile_ok(r#"
numbers = [5, 3, 8, 1, 9, 2, 7, 4, 6]
sorted = numbers.sort
puts sorted.join(', ')
"#);
}

// ── String processing ──────────────────────────────────────
#[test]
fn string_processing() {
    compile_ok(r#"
text = "Hello, World! This is Ruby."
words = text.split(' ')
upper = text.upcase
lower = text.downcase
stripped = "  hello  ".strip
puts words.length
puts upper
puts lower
puts stripped
"#);
}

// ── Iterators ──────────────────────────────────────────────
#[test]
fn iterators() {
    compile_ok(r#"
numbers = [1, 2, 3, 4, 5]
doubled = numbers.map { |n| n * 2 }
evens = numbers.select { |n| n % 2 == 0 }
sum = numbers.reduce(0) { |acc, n| acc + n }
puts doubled.join(', ')
puts evens.join(', ')
puts sum
"#);
}

// ── Exception handling ─────────────────────────────────────
#[test]
fn exception_handling() {
    compile_ok(r#"
begin
  x = 10
  raise 'something went wrong'
rescue => e
  puts e
ensure
  puts 'cleanup done'
end
"#);
}

// ── Module with class ──────────────────────────────────────
#[test]
fn module_with_class() {
    compile_ok(r#"
module Printable
  def to_string
    'printable'
  end
end

class Person
  include Printable
  attr_reader :name

  def initialize(name)
    @name = name
  end
end

person = Person.new('Alice')
puts person.name
"#);
}

// ── Hash operations ────────────────────────────────────────
#[test]
fn hash_operations() {
    compile_ok(r#"
person = {name: 'Alice', age: 30, city: 'NYC'}
puts person.keys.join(', ')
puts person.values.join(', ')
"#);
}
