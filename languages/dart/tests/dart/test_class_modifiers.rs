//! Dart 3 class modifiers: base, final, interface, sealed, and mixin class
//! semantics within a single library, including allowed subclassing patterns.

dart_cases! {
    base_class_instantiation_and_method => {
        r#"base class Vehicle {
  String kind() {
    return 'vehicle';
  }
}
void main() {
  print(Vehicle().kind());
}"#,
        ["vehicle"]
    };

    base_class_field_read_through_instance => {
        r#"base class Counter {
  int n = 7;
}
void main() {
  print(Counter().n);
}"#,
        ["7"]
    };

    base_class_extended_in_same_library => {
        r#"base class Animal {
  String speak() {
    return '...';
  }
}
class Dog extends Animal {
  @override
  String speak() {
    return 'woof';
  }
}
void main() {
  print(Dog().speak());
}"#,
        ["woof"]
    };

    base_subclass_inherits_base_field => {
        r#"base class Node {
  int depth = 1;
}
class Leaf extends Node {}
void main() {
  print(Leaf().depth);
}"#,
        ["1"]
    };

    base_subclass_constructor_passes_to_super => {
        r#"base class Pair {
  int a;
  Pair(this.a);
}
class DoublePair extends Pair {
  DoublePair(int v) : super(v * 2);
}
void main() {
  print(DoublePair(3).a);
}"#,
        ["6"]
    };

    base_class_method_called_on_subclass_reference => {
        r#"base class Shape {
  int sides() {
    return 0;
  }
}
class Square extends Shape {
  @override
  int sides() {
    return 4;
  }
}
void main() {
  Shape s = Square();
  print(s.sides());
}"#,
        ["4"]
    };

    base_class_with_static_member => {
        r#"base class Config {
  static int version = 3;
}
void main() {
  print(Config.version);
}"#,
        ["3"]
    };

    base_class_getter_overridden_in_subclass => {
        r#"base class Box {
  int get size {
    return 1;
  }
}
class Crate extends Box {
  @override
  int get size {
    return 10;
  }
}
void main() {
  print(Crate().size);
}"#,
        ["10"]
    };

    base_class_two_level_extension_chain => {
        r#"base class A {
  int val() {
    return 1;
  }
}
class B extends A {}
class C extends B {
  @override
  int val() {
    return super.val() + 2;
  }
}
void main() {
  print(C().val());
}"#,
        ["3"]
    };

    base_class_instance_method_uses_this => {
        r#"base class Tag {
  String label = 'x';
  String show() {
    return this.label;
  }
}
void main() {
  print(Tag().show());
}"#,
        ["x"]
    };

    final_class_direct_instantiation => {
        r#"final class Token {
  String value;
  Token(this.value);
}
void main() {
  print(Token('abc').value);
}"#,
        ["abc"]
    };

    final_class_method_returns_computed_value => {
        r#"final class Rect {
  int w = 4;
  int h = 5;
  int area() {
    return w * h;
  }
}
void main() {
  print(Rect().area());
}"#,
        ["20"]
    };

    final_class_field_mutation_after_construction => {
        r#"final class Acc {
  int total = 0;
  void add(int n) {
    total = total + n;
  }
}
void main() {
  var a = Acc();
  a.add(3);
  a.add(4);
  print(a.total);
}"#,
        ["7"]
    };

    final_class_with_named_constructor => {
        r#"final class Point {
  int x;
  int y;
  Point(this.x, this.y);
  Point.origin() : x = 0, y = 0;
}
void main() {
  print(Point.origin().x);
}"#,
        ["0"]
    };

    final_class_getter_without_setter => {
        r#"final class ReadOnly {
  int get count {
    return 42;
  }
}
void main() {
  print(ReadOnly().count);
}"#,
        ["42"]
    };

    final_class_static_method => {
        r#"final class MathUtil {
  static int doubleIt(int n) {
    return n * 2;
  }
}
void main() {
  print(MathUtil.doubleIt(11));
}"#,
        ["22"]
    };

    final_class_equality_two_instances => {
        r#"final class Item {
  int id;
  Item(this.id);
}
void main() {
  print(Item(1) == Item(1));
}"#,
        ["false"]
    };

    final_class_to_string_default => {
        r#"final class Widget {}
void main() {
  print(Widget().toString().contains('Widget'));
}"#,
        ["true"]
    };

    interface_class_implemented_by_concrete_class => {
        r#"interface class Drawable {
  String draw();
}
class Circle implements Drawable {
  @override
  String draw() {
    return 'circle';
  }
}
void main() {
  print(Circle().draw());
}"#,
        ["circle"]
    };

    interface_class_getter_implementation => {
        r#"interface class Sized {
  int get width;
}
class Banner implements Sized {
  @override
  int get width {
    return 320;
  }
}
void main() {
  print(Banner().width);
}"#,
        ["320"]
    };

    interface_class_multiple_implementers => {
        r#"interface class Identifiable {
  String id();
}
class User implements Identifiable {
  @override
  String id() {
    return 'u1';
  }
}
class Guest implements Identifiable {
  @override
  String id() {
    return 'g1';
  }
}
void main() {
  print(User().id() + Guest().id());
}"#,
        ["u1g1"]
    };

    interface_class_method_via_interface_typed_variable => {
        r#"interface class Runner {
  int pace();
}
class Sprinter implements Runner {
  @override
  int pace() {
    return 12;
  }
}
void main() {
  Runner r = Sprinter();
  print(r.pace());
}"#,
        ["12"]
    };

    interface_class_with_field_on_implementer => {
        r#"interface class Named {
  String name();
}
class Person implements Named {
  String _name = 'Ann';
  @override
  String name() {
    return _name;
  }
}
void main() {
  print(Person().name());
}"#,
        ["Ann"]
    };

    interface_class_implements_two_interfaces => {
        r#"interface class Readable {
  String read();
}
interface class Writable {
  void write(String s);
}
class Buffer implements Readable, Writable {
  String data = '';
  @override
  String read() {
    return data;
  }
  @override
  void write(String s) {
    data = s;
  }
}
void main() {
  var b = Buffer();
  b.write('ok');
  print(b.read());
}"#,
        ["ok"]
    };

    interface_class_setter_implementation => {
        r#"interface class Mutable {
  set value(int v);
  int get value;
}
class Cell implements Mutable {
  int _v = 0;
  @override
  int get value {
    return _v;
  }
  @override
  set value(int v) {
    _v = v;
  }
}
void main() {
  var c = Cell();
  c.value = 9;
  print(c.value);
}"#,
        ["9"]
    };

    sealed_class_direct_subtype_instantiation => {
        r#"sealed class Result {}
class Ok extends Result {
  int value;
  Ok(this.value);
}
void main() {
  print(Ok(5).value);
}"#,
        ["5"]
    };

    sealed_class_indirect_subtype_through_base => {
        r#"sealed class Expr {}
class NumLit extends Expr {
  int n;
  NumLit(this.n);
}
class AddExpr extends Expr {
  Expr left;
  Expr right;
  AddExpr(this.left, this.right);
}
void main() {
  var tree = AddExpr(NumLit(2), NumLit(3));
  print((tree.left as NumLit).n);
}"#,
        ["2"]
    };

    sealed_exhaustive_switch_on_direct_subtype => {
        r#"sealed class Status {}
class Active extends Status {}
class Inactive extends Status {}
int code(Status s) {
  switch (s) {
    case Active():
      return 1;
    case Inactive():
      return 0;
  }
}
void main() {
  print(code(Active()));
}"#,
        ["1"]
    };

    sealed_switch_returns_value_from_second_subtype => {
        r#"sealed class Color {}
class Red extends Color {}
class Blue extends Color {}
String label(Color c) {
  switch (c) {
    case Red():
      return 'r';
    case Blue():
      return 'b';
  }
}
void main() {
  print(label(Blue()));
}"#,
        ["b"]
    };

    sealed_hierarchy_three_subtypes_all_matched => {
        r#"sealed class Shape {}
class Circle extends Shape {
  int r;
  Circle(this.r);
}
class Square extends Shape {
  int side;
  Square(this.side);
}
class Triangle extends Shape {
  int base;
  Triangle(this.base);
}
int measure(Shape s) {
  switch (s) {
    case Circle(r: var radius):
      return radius;
    case Square(side: var s):
      return s;
    case Triangle(base: var b):
      return b;
  }
}
void main() {
  print(measure(Square(6)));
}"#,
        ["6"]
    };

    sealed_subtype_with_method => {
        r#"sealed class Msg {}
class TextMsg extends Msg {
  String text;
  TextMsg(this.text);
  int length() {
    return text.length;
  }
}
void main() {
  print(TextMsg('hi').length());
}"#,
        ["2"]
    };

    sealed_switch_with_object_pattern_field => {
        r#"sealed class Response {}
class Success extends Response {
  int code;
  Success(this.code);
}
class Failure extends Response {
  String reason;
  Failure(this.reason);
}
String describe(Response r) {
  switch (r) {
    case Success(code: 200):
      return 'ok';
    case Success(code: var c):
      return 'code:$c';
    case Failure(reason: var msg):
      return msg;
  }
}
void main() {
  print(describe(Success(200)));
}"#,
        ["ok"]
    };

    sealed_class_field_on_subtype => {
        r#"sealed class Event {}
class Click extends Event {
  int x;
  int y;
  Click(this.x, this.y);
}
void main() {
  var e = Click(3, 4);
  print(e.x + e.y);
}"#,
        ["7"]
    };

    mixin_class_used_with_on_host_class => {
        r#"mixin class Timestamped {
  int stamp = 1;
}
class Record with Timestamped {}
void main() {
  print(Record().stamp);
}"#,
        ["1"]
    };

    mixin_class_method_available_after_with => {
        r#"mixin class Loggable {
  String tag() {
    return 'log';
  }
}
class App with Loggable {}
void main() {
  print(App().tag());
}"#,
        ["log"]
    };

    mixin_class_direct_instantiation => {
        r#"mixin class Marker {
  int flag = 9;
}
void main() {
  print(Marker().flag);
}"#,
        ["9"]
    };

    mixin_class_with_extends_and_with => {
        r#"class Base {
  int baseVal() {
    return 1;
  }
}
mixin class Extra {
  int extraVal() {
    return 2;
  }
}
class Combined extends Base with Extra {}
void main() {
  var c = Combined();
  print(c.baseVal() + c.extraVal());
}"#,
        ["3"]
    };

    mixin_class_multiple_with_on_class => {
        r#"mixin class A {
  int a() {
    return 1;
  }
}
mixin class B {
  int b() {
    return 2;
  }
}
class Both with A, B {}
void main() {
  print(Both().a() + Both().b());
}"#,
        ["3"]
    };

    mixin_class_field_mutation_via_host => {
        r#"mixin class Counter {
  int n = 0;
  void inc() {
    n++;
  }
}
class Box with Counter {}
void main() {
  var b = Box();
  b.inc();
  b.inc();
  print(b.n);
}"#,
        ["2"]
    };

    mixin_class_on_constraint_with_supertype => {
        r#"class Animal {
  String kind = 'animal';
}
mixin class Pet on Animal {
  String care() {
    return 'feed';
  }
}
class Dog extends Animal with Pet {}
void main() {
  print(Dog().care());
}"#,
        ["feed"]
    };

    base_and_sealed_distinct_hierarchies => {
        r#"base class BaseNode {
  int depth = 1;
}
class BaseLeaf extends BaseNode {}
sealed class Expr {}
class Lit extends Expr {
  int v;
  Lit(this.v);
}
void main() {
  print(BaseLeaf().depth + Lit(2).v);
}"#,
        ["3"]
    };

    interface_and_final_independent_types => {
        r#"interface class Port {
  int number();
}
final class Endpoint implements Port {
  int _n;
  Endpoint(this._n);
  @override
  int number() {
    return _n;
  }
}
void main() {
  print(Endpoint(8080).number());
}"#,
        ["8080"]
    };

    sealed_switch_nested_in_function => {
        r#"sealed class Option {}
class Some extends Option {
  int value;
  Some(this.value);
}
class None extends Option {}
int unwrap(Option o) {
  switch (o) {
    case Some(value: var v):
      return v;
    case None():
      return -1;
  }
}
void main() {
  print(unwrap(Some(99)));
}"#,
        ["99"]
    };

    base_class_implements_interface_class => {
        r#"interface class Describable {
  String describe();
}
base class Entity implements Describable {
  @override
  String describe() {
    return 'entity';
  }
}
class User extends Entity {
  @override
  String describe() {
    return 'user';
  }
}
void main() {
  print(User().describe());
}"#,
        ["user"]
    };

    sealed_subtype_list_in_same_library => {
        r#"sealed class Token {}
class Alpha extends Token {}
class Beta extends Token {}
class Gamma extends Token {}
int rank(Token t) {
  switch (t) {
    case Alpha():
      return 1;
    case Beta():
      return 2;
    case Gamma():
      return 3;
  }
}
void main() {
  print(rank(Beta()));
}"#,
        ["2"]
    };

    mixin_class_getter_on_host => {
        r#"mixin class HasId {
  int get id {
    return 7;
  }
}
class Node with HasId {}
void main() {
  print(Node().id);
}"#,
        ["7"]
    };

    final_class_with_factory_redirect => {
        r#"final class Point {
  int x;
  int y;
  Point(this.x, this.y);
  factory Point.zero() {
    return Point(0, 0);
  }
}
void main() {
  print(Point.zero().x);
}"#,
        ["0"]
    };

    interface_class_abstract_method_only => {
        r#"interface class Service {
  int run();
}
class LocalService implements Service {
  @override
  int run() {
    return 42;
  }
}
void main() {
  print(LocalService().run());
}"#,
        ["42"]
    };

    base_class_override_calls_super => {
        r#"base class Greeter {
  String greet() {
    return 'hi';
  }
}
class LoudGreeter extends Greeter {
  @override
  String greet() {
    return super.greet() + '!';
  }
}
void main() {
  print(LoudGreeter().greet());
}"#,
        ["hi!"]
    };

    sealed_class_switch_with_guard_like_pattern => {
        r#"sealed class Num {}
class Positive extends Num {
  int v;
  Positive(this.v);
}
class Zero extends Num {}
String sign(Num n) {
  switch (n) {
    case Positive(v: var x) when x > 0:
      return 'pos';
    case Zero():
      return 'zero';
    case Positive():
      return 'other';
  }
}
void main() {
  print(sign(Positive(5)));
}"#,
        ["pos"]
    };

    mixin_class_static_method_on_class_name => {
        r#"mixin class Util {
  static int twice(int n) {
    return n * 2;
  }
}
void main() {
  print(Util.twice(6));
}"#,
        ["12"]
    };
}
