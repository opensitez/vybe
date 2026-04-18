use vybec::parser_ruby::parse;
use vybec::compiler_ruby::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

// ── Basic class ────────────────────────────────────────────
#[test] fn empty_class() { compile_ok("class Animal\nend"); }
#[test] fn class_with_method() { compile_ok("class Dog\n  def bark\n    puts 'woof'\n  end\nend"); }
#[test] fn class_initialize() { compile_ok("class Person\n  def initialize(name)\n    @name = name\n  end\nend"); }
#[test] fn class_new() { compile_ok("class Person\n  def initialize(name)\n    @name = name\n  end\nend\nperson = Person.new('Alice')"); }
#[test] fn class_method_call() { compile_ok("class Dog\n  def initialize(name)\n    @name = name\n  end\n  def bark\n    puts @name\n  end\nend\nd = Dog.new('Rex')\nd.bark"); }

// ── Inheritance ────────────────────────────────────────────
#[test] fn inheritance() { compile_ok("class Animal\n  def speak\n    puts 'generic'\n  end\nend\nclass Dog < Animal\n  def speak\n    puts 'woof'\n  end\nend"); }
#[test] fn super_call() { compile_ok("class Animal\n  def initialize(name)\n    @name = name\n  end\nend\nclass Dog < Animal\n  def initialize(name, breed)\n    super(name)\n    @breed = breed\n  end\nend"); }

// ── attr_reader / attr_writer / attr_accessor ──────────────
#[test] fn attr_reader() { compile_ok("class Person\n  attr_reader :name, :age\n  def initialize(name, age)\n    @name = name\n    @age = age\n  end\nend"); }
#[test] fn attr_accessor() { compile_ok("class Point\n  attr_accessor :x, :y\n  def initialize(x, y)\n    @x = x\n    @y = y\n  end\nend"); }
#[test] fn attr_writer() { compile_ok("class Config\n  attr_writer :debug\nend"); }

// ── Instance variables ─────────────────────────────────────
#[test] fn instance_vars() { compile_ok("class Counter\n  def initialize\n    @count = 0\n  end\n  def increment\n    @count += 1\n  end\n  def value\n    @count\n  end\nend"); }

// ── Class variables ────────────────────────────────────────
#[test] fn class_vars() { compile_ok("class Counter\n  @@total = 0\n  def initialize\n    @@total += 1\n  end\nend"); }

// ── Self methods (class methods) ───────────────────────────
#[test] fn class_method() { compile_ok("class MathUtils\n  def self.square(x)\n    x * x\n  end\nend"); }

// ── Modules ────────────────────────────────────────────────
#[test] fn module_def() { compile_ok("module Greetable\n  def greet\n    puts 'hello'\n  end\nend"); }
#[test] fn module_include() { compile_ok("module Greetable\n  def greet\n    puts 'hello'\n  end\nend\nclass Person\n  include Greetable\nend"); }

// ── Scope resolution ───────────────────────────────────────
#[test] fn scope_resolution() { compile_ok("module MyModule\n  def hello\n    'hello'\n  end\nend"); }
