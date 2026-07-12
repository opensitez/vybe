use super::helpers::{compile_ok, run_prints};

// ── implements ───────────────────────────────────────────────

#[test]
fn implements_interface() {
    compile_ok(
        "abstract class Printable { void print_(); } class Doc implements Printable { void print_() { print('doc'); } }",
    );
}

#[test]
fn implements_multiple() {
    compile_ok(
        r#"
abstract class Readable { String read(); }
abstract class Writable { void write(String s); }
class File implements Readable, Writable {
  String read() => '';
  void write(String s) {}
}
"#,
    );
}

#[test]
fn implements_with_extends() {
    compile_ok(
        r#"
abstract class Serializable { String serialize(); }
class Base { int id; Base(this.id); }
class Entity extends Base implements Serializable {
  Entity(int id) : super(id);
  String serialize() => '{"id": $id}';
}
"#,
    );
}

// ── Multiple mixins ──────────────────────────────────────────

#[test]
fn multiple_mixins() {
    compile_ok(
        r#"
mixin Flyable { void fly() { print('flying'); } }
mixin Swimmable { void swim() { print('swimming'); } }
class Duck with Flyable, Swimmable {
  String name;
  Duck(this.name);
}
"#,
    );
}

#[test]
fn mixin_on_constraint() {
    compile_ok(
        r#"
class Animal { String name; Animal(this.name); }
mixin Domestic on Animal { String get owner => 'human'; }
class Dog extends Animal with Domestic { Dog(String n) : super(n); }
"#,
    );
}

#[test]
fn mixin_with_field() {
    compile_ok(
        r#"
mixin Timestamped {
  late DateTime createdAt;
  void initTimestamp() { createdAt = DateTime.now(); }
}
class Post with Timestamped { String content; Post(this.content); }
"#,
    );
}

#[test]
fn multiple_mixins_result() {
    let out = run_prints(
        r#"
mixin Greet { String greet() => 'hello'; }
mixin Bye { String bye() => 'goodbye'; }
class Person with Greet, Bye { String name; Person(this.name); }
void main() {
  var p = Person('Alice');
  print(p.greet());
}
"#,
    );
    assert_eq!(out, ["hello"]);
}

// ── abstract classes ─────────────────────────────────────────

#[test]
fn abstract_with_concrete() {
    compile_ok(
        r#"
abstract class Vehicle {
  String name;
  Vehicle(this.name);
  void fuel() { print('fueling $name'); }
  void drive();
}
class Car extends Vehicle {
  Car(String n) : super(n);
  void drive() { print('$name driving'); }
}
"#,
    );
}

#[test]
fn abstract_static_field() {
    compile_ok(
        r#"
abstract class Config {
  static const String version = '1.0';
  String get name;
}
class AppConfig extends Config {
  String get name => 'MyApp';
}
"#,
    );
}

// ── @override ─────────────────────────────────────────────────

#[test]
fn override_method() {
    compile_ok(
        r#"
class Animal { String speak() => 'animal sound'; }
class Cat extends Animal {
  @override
  String speak() => 'meow';
}
"#,
    );
}

#[test]
fn override_getter() {
    compile_ok(
        r#"
class Base { int get value => 0; }
class Child extends Base {
  @override
  int get value => 42;
}
"#,
    );
}

#[test]
fn override_result() {
    let out = run_prints(
        r#"
class Animal { String speak() => 'generic'; }
class Dog extends Animal {
  @override
  String speak() => 'woof';
}
void main() { var d = Dog(); print(d.speak()); }
"#,
    );
    assert_eq!(out, ["woof"]);
}

// ── Redirecting constructors ─────────────────────────────────

#[test]
fn redirecting_constructor() {
    compile_ok(
        "class Point { int x; int y; Point(this.x, this.y); Point.origin() : this(0, 0); Point.unit() : this(1, 1); }",
    );
}

#[test]
fn redirecting_constructor_result() {
    let out = run_prints(
        r#"
class Point { int x; int y; Point(this.x, this.y); Point.origin() : this(0, 0); }
void main() { var p = Point.origin(); print(p.x); }
"#,
    );
    assert_eq!(out, ["0"]);
}

// ── Const constructors ───────────────────────────────────────

#[test]
fn const_constructor() {
    compile_ok("class Immutable { final int x; final int y; const Immutable(this.x, this.y); }");
}

#[test]
fn const_constructor_used() {
    compile_ok(
        "class Vec { final int x; final int y; const Vec(this.x, this.y); } const origin = Vec(0, 0);",
    );
}

// ── Static members ────────────────────────────────────────────

#[test]
fn static_counter() {
    compile_ok(
        r#"
class Counter {
  static int _count = 0;
  static void increment() { _count++; }
  static int get count => _count;
}
"#,
    );
}

#[test]
fn static_factory_pattern() {
    compile_ok(
        r#"
class Singleton {
  static Singleton? _instance;
  static Singleton get instance {
    _instance ??= Singleton._();
    return _instance!;
  }
  Singleton._();
}
"#,
    );
}

#[test]
fn static_result() {
    let out = run_prints(
        r#"
class Counter {
  static int count = 0;
  static void inc() { count++; }
}
void main() { Counter.inc(); Counter.inc(); print(Counter.count); }
"#,
    );
    assert_eq!(out, ["2"]);
}

// ── toString override ────────────────────────────────────────

#[test]
fn to_string_override() {
    compile_ok(
        "class Point { int x; int y; Point(this.x, this.y); String toString() => 'Point($x, $y)'; }",
    );
}

#[test]
fn to_string_result() {
    let out = run_prints(
        r#"
class Point { int x; int y; Point(this.x, this.y); String toString() => '($x, $y)'; }
void main() { var p = Point(3, 4); print(p); }
"#,
    );
    assert_eq!(out, ["(3, 4)"]);
}

// ── hashCode and == override ─────────────────────────────────

#[test]
fn equality_override() {
    compile_ok(
        r#"
class Point {
  int x; int y;
  Point(this.x, this.y);
  bool operator ==(Object other) {
    if (other is Point) return x == other.x && y == other.y;
    return false;
  }
  int get hashCode => x * 31 + y;
}
"#,
    );
}

// ── Deep inheritance ─────────────────────────────────────────

#[test]
fn three_level_inheritance() {
    compile_ok(
        r#"
class A { String name() => 'A'; }
class B extends A { String extra() => 'B'; }
class C extends B { String more() => 'C'; }
void main() { var c = C(); print(c.name()); }
"#,
    );
}

#[test]
fn three_level_result() {
    let out = run_prints(
        r#"
class A { int val() => 1; }
class B extends A { int bonus() => 10; }
class C extends B {}
void main() { var c = C(); print(c.val() + c.bonus()); }
"#,
    );
    assert_eq!(out, ["11"]);
}

// ── Mixin method resolution order ────────────────────────────

#[test]
fn mixin_overrides_base() {
    let out = run_prints(
        r#"
class Base { String greet() => 'base'; }
mixin Override { String greet() => 'mixin'; }
class Child extends Base with Override {}
void main() { print(Child().greet()); }
"#,
    );
    assert_eq!(out, ["mixin"]);
}

// ── Private members ──────────────────────────────────────────

#[test]
fn private_field() {
    compile_ok(
        "class Account { double _balance = 0; void deposit(double amt) { _balance += amt; } double get balance => _balance; }",
    );
}

#[test]
fn private_method() {
    compile_ok(
        "class Parser { String _clean(String s) => s.trim(); String parse(String s) => _clean(s); }",
    );
}

#[test]
fn private_field_result() {
    let out = run_prints(
        r#"
class Account { double _balance = 0; void deposit(double v) { _balance += v; } double get balance => _balance; }
void main() { var a = Account(); a.deposit(100); print(a.balance); }
"#,
    );
    assert_eq!(out, ["100"]);
}

// ── Extension methods (more cases) ───────────────────────────

#[test]
fn extension_on_int() {
    compile_ok("extension IntExt on int { bool get isEven => this % 2 == 0; }");
}

#[test]
fn extension_on_list() {
    compile_ok(
        "extension ListExt<T> on List<T> { T? get secondOrNull => length > 1 ? this[1] : null; }",
    );
}

#[test]
fn extension_chain() {
    compile_ok(
        r#"
extension StringX on String {
  String shout() => toUpperCase() + '!';
}
void main() { print('hello'.shout()); }
"#,
    );
}
