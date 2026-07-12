use super::helpers::{compile_ok, run_ruby};

// ── Class definitions (compile) ─────────────────────────────────────────────

#[test]
fn class_empty() {
    compile_ok("class Foo\nend\n");
}
#[test]
fn class_initialize() {
    compile_ok("class Dog\n  def initialize(name)\n    @name = name\n  end\nend\n");
}
#[test]
fn class_method() {
    compile_ok("class Dog\n  def bark\n    puts 'woof'\n  end\nend\n");
}
#[test]
fn class_inherit() {
    compile_ok("class Animal\nend\nclass Dog < Animal\nend\n");
}
#[test]
fn class_attr_reader() {
    compile_ok(
        "class Dog\n  attr_reader :name\n  def initialize(name)\n    @name = name\n  end\nend\n",
    );
}
#[test]
fn class_attr_writer() {
    compile_ok("class Dog\n  attr_writer :name\nend\n");
}
#[test]
fn class_attr_accessor() {
    compile_ok("class Dog\n  attr_accessor :name, :age\nend\n");
}
#[test]
fn class_self_method() {
    compile_ok("class Util\n  def self.hello\n    puts 'hi'\n  end\nend\n");
}
#[test]
fn class_class_var() {
    compile_ok("class Counter\n  @@count = 0\n  def initialize\n    @@count += 1\n  end\nend\n");
}
#[test]
fn class_super() {
    compile_ok(
        "class Animal\n  def speak\n    'generic'\n  end\nend\nclass Dog < Animal\n  def speak\n    super\n  end\nend\n",
    );
}

// ── Module ──────────────────────────────────────────────────────────────────

#[test]
fn module_def() {
    compile_ok("module Greetable\n  def greet\n    puts 'hello'\n  end\nend\n");
}
#[test]
fn module_include() {
    compile_ok(
        "module Greetable\n  def greet\n    puts 'hello'\n  end\nend\nclass Person\n  include Greetable\nend\n",
    );
}

// ── Runtime ─────────────────────────────────────────────────────────────────

#[test]
fn class_create_instance() {
    let out = run_ruby(
        "class Dog\n  def initialize(name)\n    @name = name\n  end\n  def bark\n    puts 'woof'\n  end\nend\nd = Dog.new('Rex')\nd.bark\n",
    );
    assert_eq!(out, vec!["woof"]);
}

#[test]
fn class_attr_reader_runtime() {
    let out = run_ruby(
        "class Dog\n  attr_reader :name\n  def initialize(name)\n    @name = name\n  end\nend\nd = Dog.new('Rex')\nputs d.name\n",
    );
    assert_eq!(out, vec!["Rex"]);
}

#[test]
fn class_attr_accessor_runtime() {
    let out = run_ruby(
        "class Dog\n  attr_accessor :name\n  def initialize(name)\n    @name = name\n  end\nend\nd = Dog.new('Rex')\nd.name = 'Buddy'\nputs d.name\n",
    );
    assert_eq!(out, vec!["Buddy"]);
}

#[test]
fn class_inherit_runtime() {
    let out = run_ruby(
        "class Animal\n  def speak\n    puts 'generic'\n  end\nend\nclass Dog < Animal\n  def speak\n    puts 'woof'\n  end\nend\nd = Dog.new\nd.speak\n",
    );
    assert_eq!(out, vec!["woof"]);
}

#[test]
fn class_self_method_runtime() {
    let out = run_ruby("class Util\n  def self.hello\n    puts 'hi'\n  end\nend\nUtil.hello\n");
    assert_eq!(out, vec!["hi"]);
}

#[test]
fn class_instance_vars_runtime() {
    let out = run_ruby(
        "class Point\n  def initialize(x, y)\n    @x = x\n    @y = y\n  end\n  def to_s\n    puts @x\n    puts @y\n  end\nend\np = Point.new(3, 4)\np.to_s\n",
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn class_super_minimal_runtime() {
    // Minimal super() test
    let out = run_ruby(
        "class A\n  def initialize(x)\n    @x = x\n  end\n  def show\n    puts @x\n  end\nend\nclass B < A\n  def initialize\n    super(42)\n  end\nend\nb = B.new\nb.show\n",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn class_super_with_child_method_runtime() {
    let out = run_ruby(
        "class A\n  def initialize(x)\n    @x = x\n  end\nend\nclass B < A\n  def initialize\n    super(42)\n  end\n  def greet\n    puts 'hello'\n  end\nend\nb = B.new\nb.greet\n",
    );
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn class_method_return_value_runtime() {
    // Method that returns a value via implicit return
    let out = run_ruby(
        "class Foo\n  def initialize(x)\n    @x = x\n  end\n  def val\n    @x\n  end\nend\nf = Foo.new(42)\nputs f.val\n",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn class_method_return_expr_runtime() {
    // Method that returns a binary expression via implicit return
    let out = run_ruby(
        "class Foo\n  def initialize(x)\n    @x = x\n  end\n  def double\n    @x * 2\n  end\nend\nf = Foo.new(5)\nputs f.double\n",
    );
    assert_eq!(out, vec!["10"]);
}
