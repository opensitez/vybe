//! Core Dart class semantics: fields, methods, constructors, this, static,
//! inheritance, super, and override.

dart_cases! {
    instance_field_reads_initial_value => {
        r#"class Box {
  int size = 10;
}
void main() {
  var b = Box();
  print(b.size);
}"#,
        ["10"]
    };

    typed_instance_field_stores_string => {
        r#"class Label {
  String text = 'hello';
}
void main() {
  var l = Label();
  print(l.text);
}"#,
        ["hello"]
    };

    final_field_set_at_construction => {
        r#"class Token {
  final String value;
  Token(this.value);
}
void main() {
  var t = Token('abc');
  print(t.value);
}"#,
        ["abc"]
    };

    field_default_initializer_applies => {
        r#"class Counter {
  int count = 0;
}
void main() {
  var c = Counter();
  print(c.count);
}"#,
        ["0"]
    };

    multiple_instance_fields_are_independent => {
        r#"class Point {
  int x = 1;
  int y = 2;
}
void main() {
  var p = Point();
  print(p.x + p.y);
}"#,
        ["3"]
    };

    instance_method_returns_computed_value => {
        r#"class Rect {
  int w = 4;
  int h = 5;
  int area() {
    return w * h;
  }
}
void main() {
  var r = Rect();
  print(r.area());
}"#,
        ["20"]
    };

    instance_method_mutates_field => {
        r#"class Acc {
  int total = 0;
  void add(int n) {
    total = total + n;
  }
}
void main() {
  var a = Acc();
  a.add(7);
  print(a.total);
}"#,
        ["7"]
    };

    method_uses_this_to_disambiguate_field => {
        r#"class Wrap {
  int value = 3;
  int bump() {
    return this.value + 1;
  }
}
void main() {
  print(Wrap().bump());
}"#,
        ["4"]
    };

    method_accepts_parameters => {
        r#"class Adder {
  int sum(int a, int b) {
    return a + b;
  }
}
void main() {
  print(Adder().sum(6, 7));
}"#,
        ["13"]
    };

    arrow_body_method_returns_expression => {
        r#"class Twice {
  int go(int n) => n * 2;
}
void main() {
  print(Twice().go(11));
}"#,
        ["22"]
    };

    this_field_constructor_shorthand => {
        r#"class Pair {
  int a;
  int b;
  Pair(this.a, this.b);
}
void main() {
  var p = Pair(2, 3);
  print(p.a + p.b);
}"#,
        ["5"]
    };

    constructor_body_assigns_fields => {
        r#"class User {
  String name;
  User(String n) {
    name = n;
  }
}
void main() {
  print(User('Ann').name);
}"#,
        ["Ann"]
    };

    no_arg_constructor_creates_instance => {
        r#"class Empty {}
void main() {
  var e = Empty();
  print(e != null);
}"#,
        ["true"]
    };

    this_passed_to_other_method => {
        r#"class Node {
  int tag = 1;
  int readTag(Node other) {
    return other.tag;
  }
}
void main() {
  var n = Node();
  print(n.readTag(n));
}"#,
        ["1"]
    };

    static_field_read_via_class_name => {
        r#"class Config {
  static int version = 2;
}
void main() {
  print(Config.version);
}"#,
        ["2"]
    };

    static_method_computes_from_arguments => {
        r#"class MathUtil {
  static int max(int a, int b) {
    return a > b ? a : b;
  }
}
void main() {
  print(MathUtil.max(8, 3));
}"#,
        ["8"]
    };

    static_getter_exposes_value => {
        r#"class App {
  static int _build = 100;
  static int get build => _build;
}
void main() {
  print(App.build);
}"#,
        ["100"]
    };

    static_field_mutation_persists => {
        r#"class Hits {
  static int count = 0;
  static void bump() {
    count = count + 1;
  }
}
void main() {
  Hits.bump();
  Hits.bump();
  print(Hits.count);
}"#,
        ["2"]
    };

    static_method_invoked_without_instance => {
        r#"class IdGen {
  static int next() {
    return 42;
  }
}
void main() {
  print(IdGen.next());
}"#,
        ["42"]
    };

    extends_inherits_instance_method => {
        r#"class Base {
  int value() {
    return 5;
  }
}
class Child extends Base {}
void main() {
  print(Child().value());
}"#,
        ["5"]
    };

    extends_inherits_field_through_method => {
        r#"class Animal {
  String name = 'cat';
  String label() {
    return name;
  }
}
class Kitten extends Animal {}
void main() {
  print(Kitten().label());
}"#,
        ["cat"]
    };

    super_constructor_passes_argument_to_base => {
        r#"class Base {
  int n;
  Base(this.n);
}
class Derived extends Base {
  Derived(int v) : super(v);
}
void main() {
  print(Derived(9).n);
}"#,
        ["9"]
    };

    super_method_call_extends_base_result => {
        r#"class A {
  String greet() {
    return 'hi';
  }
}
class B extends A {
  String greet() {
    return super.greet() + '!';
  }
}
void main() {
  print(B().greet());
}"#,
        ["hi!"]
    };

    two_level_inheritance_reuses_grandparent_method => {
        r#"class A {
  int base() {
    return 1;
  }
}
class B extends A {}
class C extends B {}
void main() {
  print(C().base());
}"#,
        ["1"]
    };

    subclass_adds_method_alongside_inherited => {
        r#"class Base {
  int one() {
    return 1;
  }
}
class Sub extends Base {
  int two() {
    return 2;
  }
}
void main() {
  var s = Sub();
  print(s.one() + s.two());
}"#,
        ["3"]
    };

    override_method_replaces_superclass_behavior => {
        r#"class Animal {
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

    override_getter_replaces_superclass_property => {
        r#"class Base {
  int get val {
    return 0;
  }
}
class Child extends Base {
  @override
  int get val {
    return 99;
  }
}
void main() {
  print(Child().val);
}"#,
        ["99"]
    };

    override_can_call_super_implementation => {
        r#"class Base {
  int calc() {
    return 10;
  }
}
class Child extends Base {
  @override
  int calc() {
    return super.calc() + 5;
  }
}
void main() {
  print(Child().calc());
}"#,
        ["15"]
    };

    late_field_initialized_before_read => {
        r#"class Lazy {
  late int value;
  Lazy() {
    value = 7;
  }
}
void main() {
  print(Lazy().value);
}"#,
        ["7"]
    };

    private_field_exposed_via_public_getter => {
        r#"class Vault {
  int _secret = 42;
  int get secret {
    return _secret;
  }
}
void main() {
  print(Vault().secret);
}"#,
        ["42"]
    };

    instance_method_calls_sibling_method => {
        r#"class Calc {
  int doubleIt(int n) {
    return n * 2;
  }
  int quadruple(int n) {
    return doubleIt(doubleIt(n));
  }
}
void main() {
  print(Calc().quadruple(3));
}"#,
        ["12"]
    };

    static_method_calls_other_static_method => {
        r#"class Chain {
  static int step1(int n) {
    return n + 1;
  }
  static int step2(int n) {
    return step1(n) * 2;
  }
}
void main() {
  print(Chain.step2(4));
}"#,
        ["10"]
    };

    instance_method_reads_static_field => {
        r#"class Reader {
  static int shared = 6;
  int readShared() {
    return shared;
  }
}
void main() {
  print(Reader().readShared());
}"#,
        ["6"]
    };

    getter_returns_computed_value_from_fields => {
        r#"class Rect {
  int w = 3;
  int h = 4;
  int get area {
    return w * h;
  }
}
void main() {
  print(Rect().area);
}"#,
        ["12"]
    };

    setter_updates_backing_field => {
        r#"class Box {
  int _size = 1;
  int get size {
    return _size;
  }
  set size(int v) {
    _size = v;
  }
}
void main() {
  var b = Box();
  b.size = 8;
  print(b.size);
}"#,
        ["8"]
    };

    to_string_override_formats_instance => {
        r#"class Point {
  int x = 1;
  int y = 2;
  String toString() {
    return '($x,$y)';
  }
}
void main() {
  print(Point());
}"#,
        ["(1,2)"]
    };

    constructor_with_default_parameter_value => {
        r#"class Greeter {
  String msg;
  Greeter([this.msg = 'hi']) {}
}
void main() {
  print(Greeter().msg);
}"#,
        ["hi"]
    };

    method_returns_this_for_chaining => {
        r#"class Builder {
  int v = 0;
  Builder add(int n) {
    v = v + n;
    return this;
  }
}
void main() {
  var b = Builder();
  b.add(2).add(3);
  print(b.v);
}"#,
        ["5"]
    };

    subclass_constructor_sets_own_field => {
        r#"class Base {
  int a = 0;
}
class Sub extends Base {
  int b = 0;
  Sub(int x) {
    b = x;
  }
}
void main() {
  print(Sub(4).b);
}"#,
        ["4"]
    };

    static_field_not_shared_with_instance_name => {
        r#"class Demo {
  static int tag = 1;
  int readTag() {
    return tag;
  }
}
void main() {
  print(Demo().readTag());
}"#,
        ["1"]
    };

    override_setter_updates_subclass_storage => {
        r#"class Base {
  int _v = 0;
  set val(int n) {
    _v = n;
  }
  int get val {
    return _v;
  }
}
class Child extends Base {
  @override
  set val(int n) {
    _v = n * 2;
  }
}
void main() {
  var c = Child();
  c.val = 3;
  print(c.val);
}"#,
        ["6"]
    };

    inherited_method_visible_on_subclass_reference => {
        r#"class Parent {
  String id() {
    return 'p';
  }
}
class Kid extends Parent {}
void main() {
  Parent p = Kid();
  print(p.id());
}"#,
        ["p"]
    };

    field_mutation_visible_to_other_methods => {
        r#"class State {
  int n = 0;
  void inc() {
    n = n + 1;
  }
  int read() {
    return n;
  }
}
void main() {
  var s = State();
  s.inc();
  s.inc();
  print(s.read());
}"#,
        ["2"]
    };

    method_with_named_optional_parameter => {
        r#"class Greeter {
  String build({String name = 'world'}) {
    return 'hello $name';
  }
}
void main() {
  print(Greeter().build());
}"#,
        ["hello world"]
    };

    super_used_in_constructor_initializer_with_body => {
        r#"class Base {
  int n;
  Base(this.n);
}
class Sub extends Base {
  int m;
  Sub(int a, int b) : super(a) {
    m = b;
  }
}
void main() {
  var s = Sub(2, 3);
  print(s.n + s.m);
}"#,
        ["5"]
    };

    static_and_instance_methods_coexist => {
        r#"class Mix {
  int inst() {
    return 1;
  }
  static int stat() {
    return 2;
  }
}
void main() {
  print(Mix().inst() + Mix.stat());
}"#,
        ["3"]
    };

    subclass_overrides_method_parent_still_unaffected => {
        r#"class Base {
  String tag() {
    return 'base';
  }
}
class Sub extends Base {
  @override
  String tag() {
    return 'sub';
  }
}
void main() {
  print(Base().tag());
}"#,
        ["base"]
    };

    constructor_initializes_multiple_this_fields => {
        r#"class Triple {
  int a;
  int b;
  int c;
  Triple(this.a, this.b, this.c);
}
void main() {
  var t = Triple(1, 2, 3);
  print(t.a + t.b + t.c);
}"#,
        ["6"]
    };

    instance_equality_reference_not_same_for_two_instances => {
        r#"class Item {}
void main() {
  var a = Item();
  var b = Item();
  print(a == b);
}"#,
        ["false"]
    };

    class_method_returning_string_concatenation => {
        r#"class Name {
  String first = 'Ada';
  String last = 'Lovelace';
  String full() {
    return first + ' ' + last;
  }
}
void main() {
  print(Name().full());
}"#,
        ["Ada Lovelace"]
    };
}
