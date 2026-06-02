use super::helpers::{compile_ok, run_ruby, run_ruby_one};

#[test]
fn attr_accessor_multiple() {
    compile_ok(
        "class Person\n  attr_accessor :name, :age, :email\n  def initialize(name, age, email)\n    @name = name\n    @age = age\n    @email = email\n  end\nend\n",
    );
}

#[test]
fn attr_reader_multiple() {
    compile_ok(
        "class Point\n  attr_reader :x, :y, :z\n  def initialize(x, y, z)\n    @x = x\n    @y = y\n    @z = z\n  end\nend\n",
    );
}

#[test]
fn protected_method_in_subclass() {
    compile_ok(
        "class Animal\n  protected\n  def secret\n    'hidden'\n  end\nend\nclass Dog < Animal\n  def reveal\n    secret\n  end\nend\nd = Dog.new\nd.reveal\n",
    );
}

#[test]
fn private_method_definition() {
    compile_ok(
        "class Vault\n  def open\n    unlock\n  end\n  private\n  def unlock\n    'unlocked'\n  end\nend\n",
    );
}

#[test]
fn public_send_method() {
    compile_ok(
        "class Foo\n  def greet\n    'hello'\n  end\nend\nf = Foo.new\nf.public_send(:greet)\n",
    );
}

#[test]
fn respond_to_include_private() {
    compile_ok(
        "class Vault\n  private\n  def secret\n    'hidden'\n  end\nend\nv = Vault.new\nv.respond_to?(:secret, true)\n",
    );
}

#[test]
fn class_method_self_prefix() {
    compile_ok(
        "class MathHelper\n  def self.square(n)\n    n * n\n  end\nend\nresult = MathHelper.square(7)\n",
    );
}

#[test]
fn class_method_calls_instance_via_new() {
    compile_ok(
        "class Builder\n  def build\n    'built'\n  end\n  def self.create_and_build\n    Builder.new.build\n  end\nend\nBuilder.create_and_build\n",
    );
}

#[test]
fn instance_class_method_returns_name() {
    let out = run_ruby("class Cat\nend\nc = Cat.new\nputs c.class\n");
    assert_eq!(out, vec!["Cat"]);
}

#[test]
fn is_a_check_on_instance() {
    let out = run_ruby("class Dog\nend\nd = Dog.new\nputs d.is_a?(Dog)\n");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn kind_of_alias_check() {
    let out = run_ruby("class Cat\nend\nc = Cat.new\nputs c.kind_of?(Cat)\n");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instance_of_exact_check() {
    let out = run_ruby(
        "class Animal\nend\nclass Dog < Animal\nend\nd = Dog.new\nputs d.instance_of?(Dog)\nputs d.instance_of?(Animal)\n",
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn nil_check_on_instance_and_nil() {
    let out = run_ruby("class Foo\nend\nf = Foo.new\nputs f.nil?\nputs nil.nil?\n");
    assert_eq!(out, vec!["false", "true"]);
}

#[test]
fn frozen_check_unfrozen() {
    compile_ok("class Box\n  attr_accessor :value\nend\nb = Box.new\nb.frozen?\n");
}

#[test]
fn freeze_then_frozen_true() {
    compile_ok(
        "class Token\n  attr_reader :val\n  def initialize(v)\n    @val = v\n  end\nend\nt = Token.new('abc')\nt.freeze\nt.frozen?\n",
    );
}

#[test]
fn dup_creates_unfrozen_copy() {
    compile_ok(
        "class Config\n  attr_accessor :setting\nend\norig = Config.new\norig.freeze\ncopy = orig.dup\ncopy.setting = 'new'\n",
    );
}

#[test]
fn clone_copies_frozen_state() {
    compile_ok(
        "class Tag\n  attr_reader :name\n  def initialize(n)\n    @name = n\n  end\nend\nt = Tag.new('x')\nt.freeze\nc = t.clone\nc.frozen?\n",
    );
}

#[test]
fn object_id_returns_integer() {
    compile_ok("class Widget\nend\nw = Widget.new\nid = w.object_id\n");
}

#[test]
fn eq_operator_override() {
    let out = run_ruby(
        "class Money\n  def initialize(amount)\n    @amount = amount\n  end\n  def ==(other)\n    @amount == other.amount\n  end\n  def amount\n    @amount\n  end\nend\na = Money.new(10)\nb = Money.new(10)\nputs a == b\n",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn eql_override() {
    compile_ok(
        "class Vector\n  def initialize(x, y)\n    @x = x\n    @y = y\n  end\n  def eql?(other)\n    @x == other.instance_variable_get(:@x) && @y == other.instance_variable_get(:@y)\n  end\nend\nv1 = Vector.new(1, 2)\nv2 = Vector.new(1, 2)\nv1.eql?(v2)\n",
    );
}

#[test]
fn hash_override() {
    compile_ok(
        "class Key\n  def initialize(val)\n    @val = val\n  end\n  def hash\n    @val.hash\n  end\nend\nk = Key.new(42)\nk.hash\n",
    );
}

#[test]
fn spaceship_for_sorting() {
    compile_ok(
        "class Weight\n  def initialize(kg)\n    @kg = kg\n  end\n  def <=>(other)\n    @kg <=> other.instance_variable_get(:@kg)\n  end\nend\nweights = [Weight.new(5), Weight.new(2), Weight.new(8)]\nweights.sort { |a, b| a <=> b }\n",
    );
}

#[test]
fn to_s_override_in_class() {
    let out = run_ruby(
        "class Greeting\n  def initialize(msg)\n    @msg = msg\n  end\n  def to_s\n    'Greeting: ' + @msg\n  end\nend\ng = Greeting.new('hello')\nputs g.to_s\n",
    );
    assert_eq!(out, vec!["Greeting: hello"]);
}

#[test]
fn inspect_override_in_class() {
    compile_ok(
        "class Node\n  def initialize(val)\n    @val = val\n  end\n  def inspect\n    'Node(' + @val.to_s + ')'\n  end\nend\nn = Node.new(7)\nn.inspect\n",
    );
}

#[test]
fn method_alias_keyword() {
    compile_ok(
        "class Talker\n  def speak\n    'speaking'\n  end\n  alias say speak\nend\nt = Talker.new\nt.say\n",
    );
}

#[test]
fn method_alias_method_call() {
    compile_ok(
        "class Printer\n  def print_text\n    'printing'\n  end\n  alias_method :display, :print_text\nend\np = Printer.new\np.display\n",
    );
}

#[test]
fn method_missing_catch_all() {
    compile_ok(
        "class Ghost\n  def method_missing(name, *args)\n    'called ' + name.to_s\n  end\nend\ng = Ghost.new\ng.anything\n",
    );
}

#[test]
fn define_method_dynamic() {
    compile_ok(
        "class Greeter\n  ['hello', 'goodbye'].each do |word|\n    define_method(word) do\n      puts word\n    end\n  end\nend\n",
    );
}

#[test]
fn class_eval_add_method() {
    compile_ok(
        "class Robot\nend\nRobot.class_eval do\n  def beep\n    'beep'\n  end\nend\nRobot.new.beep\n",
    );
}

#[test]
fn singleton_method_on_object() {
    compile_ok("obj = Object.new\ndef obj.greet\n  'hello from singleton'\nend\nobj.greet\n");
}

#[test]
fn self_class_inside_instance_method() {
    let out = run_ruby(
        "class Tiger\n  def my_class\n    self.class\n  end\nend\nt = Tiger.new\nputs t.my_class\n",
    );
    assert_eq!(out, vec!["Tiger"]);
}

#[test]
fn multi_level_inheritance_super() {
    let out = run_ruby(
        "class A\n  def greet\n    'A'\n  end\nend\nclass B < A\n  def greet\n    super + 'B'\n  end\nend\nclass C < B\n  def greet\n    super + 'C'\n  end\nend\nc = C.new\nputs c.greet\n",
    );
    assert_eq!(out, vec!["ABC"]);
}
