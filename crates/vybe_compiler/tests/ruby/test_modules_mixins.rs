use super::helpers::compile_ok;

// ── module include ────────────────────────────────────────────
#[test]
fn module_include_adds_instance_methods() {
    compile_ok(
        r#"
module Greetable
  def greet; "Hello, I am #{name}"; end
end
class Person
  include Greetable
  attr_reader :name
  def initialize(n); @name = n; end
end
puts Person.new("Alice").greet
"#,
    );
}

#[test]
fn module_include_multiple() {
    compile_ok(
        r#"
module Swimmable
  def swim; "splashing"; end
end
module Flyable
  def fly; "soaring"; end
end
class Duck
  include Swimmable
  include Flyable
end
d = Duck.new
puts d.swim
puts d.fly
"#,
    );
}

// ── module extend ─────────────────────────────────────────────
#[test]
fn module_extend_adds_class_methods() {
    compile_ok(
        r#"
module ClassMethods
  def create(name); new(name); end
end
class Robot
  extend ClassMethods
  attr_reader :name
  def initialize(n); @name = n; end
end
puts Robot.create("R2D2").name
"#,
    );
}

#[test]
fn object_extend_adds_singleton_methods() {
    compile_ok(
        r#"
module Debuggable
  def debug_info; self.class.to_s + ': ' + inspect; end
end
obj = Object.new
obj.extend(Debuggable)
puts obj.respond_to?(:debug_info)
"#,
    );
}

// ── module_function ───────────────────────────────────────────
#[test]
fn module_function_callable_as_module_method() {
    compile_ok(
        r#"
module MathUtils
  module_function
  def square(x); x * x; end
  def cube(x); x ** 3; end
end
puts MathUtils.square(4)
puts MathUtils.cube(3)
"#,
    );
}

// ── module constants ──────────────────────────────────────────
#[test]
fn module_constant_access() {
    compile_ok(
        r#"
module Config
  VERSION = "1.0.0"
  MAX_RETRIES = 3
end
puts Config::VERSION
puts Config::MAX_RETRIES
"#,
    );
}

#[test]
fn nested_module_constant() {
    compile_ok(
        r#"
module Outer
  module Inner
    VALUE = 42
  end
end
puts Outer::Inner::VALUE
"#,
    );
}

// ── ancestors chain ───────────────────────────────────────────
#[test]
fn ancestors_includes_modules() {
    compile_ok(
        r#"
module Printable; end
module Serializable; end
class Document
  include Printable
  include Serializable
end
puts Document.ancestors.include?(Printable)
puts Document.ancestors.include?(Serializable)
"#,
    );
}

// ── prepend ───────────────────────────────────────────────────
#[test]
fn prepend_intercepts_method() {
    compile_ok(
        r#"
module Logging
  def compute(x)
    result = super
    result
  end
end
class Calculator
  prepend Logging
  def compute(x); x * 2; end
end
puts Calculator.new.compute(5)
"#,
    );
}

// ── mixin with class methods via self ─────────────────────────
#[test]
fn mixin_instance_and_class_methods() {
    compile_ok(
        r#"
module Persistable
  def self.included(base)
    base.extend(ClassMethods)
  end
  module ClassMethods
    def find(id); "Record #{id}"; end
  end
  def save; "saved"; end
end
class User
  include Persistable
end
puts User.find(1)
puts User.new.save
"#,
    );
}

// ── module used as namespace ──────────────────────────────────
#[test]
fn module_as_namespace_for_classes() {
    compile_ok(
        r#"
module Animals
  class Dog
    def speak; "woof"; end
  end
  class Cat
    def speak; "meow"; end
  end
end
puts Animals::Dog.new.speak
puts Animals::Cat.new.speak
"#,
    );
}

// ── Comparable mixin ──────────────────────────────────────────
#[test]
fn comparable_mixin_provides_operators() {
    compile_ok(
        r#"
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
"#,
    );
}

// ── Enumerable mixin ──────────────────────────────────────────
#[test]
fn enumerable_mixin_provides_reduce() {
    compile_ok(
        r#"
class NumberSet
  include Enumerable
  def initialize(*ns); @ns = ns; end
  def each(&b); @ns.each(&b); end
end
puts NumberSet.new(1, 2, 3, 4, 5).reduce(:+)
"#,
    );
}

#[test]
fn enumerable_mixin_provides_sort() {
    compile_ok(
        r#"
class NameList
  include Enumerable
  def initialize(*names); @names = names; end
  def each(&b); @names.each(&b); end
end
puts NameList.new("Charlie", "Alice", "Bob").sort.inspect
"#,
    );
}

// ── Forwardable ───────────────────────────────────────────────
#[test]
fn forwardable_delegates_methods() {
    compile_ok(
        r#"
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
"#,
    );
}

// ── module reopening ──────────────────────────────────────────
#[test]
fn module_can_be_reopened_and_extended() {
    compile_ok(
        r#"
module Formatter
  def shout; upcase + "!"; end
end
module Formatter
  def whisper; downcase + "..."; end
end
class String
  include Formatter
end
puts "hello".shout
puts "HELLO".whisper
"#,
    );
}

// ── included hook ─────────────────────────────────────────────
#[test]
fn included_callback_fires_on_include() {
    compile_ok(
        r#"
module Hookable
  def self.included(base)
    base.instance_variable_set(:@hooked, true)
  end
end
class Target
  include Hookable
end
puts Target.instance_variable_get(:@hooked)
"#,
    );
}

// ── module? and instance methods ─────────────────────────────
#[test]
fn module_instance_methods_list() {
    compile_ok(
        r#"
module Tools
  def hammer; "bang"; end
  def screwdriver; "turn"; end
end
puts Tools.instance_methods.include?(:hammer)
"#,
    );
}

// ── mixins provide default implementations ────────────────────
#[test]
fn mixin_default_implementation_overridable() {
    compile_ok(
        r#"
module Printable
  def to_print; self.class.to_s + ': default'; end
end
class Report
  include Printable
  def to_print; 'Report: custom'; end
end
class Invoice
  include Printable
end
puts Report.new.to_print
puts Invoice.new.to_print
"#,
    );
}

// ── multiple inheritance resolution (MRO) ────────────────────
#[test]
fn method_resolution_order_left_to_right() {
    compile_ok(
        r#"
module A
  def hello; "A"; end
end
module B
  def hello; "B"; end
end
class C
  include A
  include B
end
puts C.new.hello
"#,
    );
}

// ── module extend on instance ─────────────────────────────────
#[test]
fn extend_on_specific_instance() {
    compile_ok(
        r#"
module Serializable
  def serialize; "data"; end
end
obj1 = Object.new
obj2 = Object.new
obj1.extend(Serializable)
puts obj1.respond_to?(:serialize)
puts obj2.respond_to?(:serialize)
"#,
    );
}

// ── module with attr_accessor ─────────────────────────────────
#[test]
fn module_defines_attr_accessor_for_includers() {
    compile_ok(
        r#"
module HasName
  attr_accessor :name
end
class Product
  include HasName
end
p = Product.new
p.name = "Widget"
puts p.name
"#,
    );
}

// ── module include chain ──────────────────────────────────────
#[test]
fn module_include_chain_calls_super() {
    compile_ok(
        r#"
module Base
  def info; "base"; end
end
module Derived
  include Base
  def info; super + "+derived"; end
end
class Thing
  include Derived
end
puts Thing.new.info
"#,
    );
}
