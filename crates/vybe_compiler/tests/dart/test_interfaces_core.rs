//! Dart interfaces via implements: abstract classes, abstract methods,
//! single and multiple interface implementation.

dart_cases! {
    implements_single_interface_method => {
        r#"abstract class Drawable {
  String draw();
}
class Circle implements Drawable {
  String draw() {
    return 'circle';
  }
}
void main() {
  print(Circle().draw());
}"#,
        ["circle"]
    };

    implements_multiple_interfaces => {
        r#"abstract class Readable {
  String read();
}
abstract class Writable {
  void write(String s);
}
class Buffer implements Readable, Writable {
  String data = '';
  String read() {
    return data;
  }
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

    abstract_method_must_be_implemented => {
        r#"abstract class Compute {
  int eval();
}
class Add implements Compute {
  int eval() {
    return 2 + 3;
  }
}
void main() {
  print(Add().eval());
}"#,
        ["5"]
    };

    implements_with_concrete_fields => {
        r#"abstract class Identifiable {
  String id();
}
class User implements Identifiable {
  String name = 'Ann';
  String id() {
    return name;
  }
}
void main() {
  print(User().id());
}"#,
        ["Ann"]
    };

    extends_and_implements_together => {
        r#"class Base {
  int baseVal() {
    return 1;
  }
}
abstract class Extra {
  int extraVal();
}
class Both extends Base implements Extra {
  int extraVal() {
    return 2;
  }
}
void main() {
  var b = Both();
  print(b.baseVal() + b.extraVal());
}"#,
        ["3"]
    };

    interface_getter_implementation => {
        r#"abstract class HasSize {
  int get size;
}
class Box implements HasSize {
  int _s = 4;
  int get size {
    return _s;
  }
}
void main() {
  print(Box().size);
}"#,
        ["4"]
    };

    interface_setter_implementation => {
        r#"abstract class Mutable {
  set value(int v);
  int get value;
}
class Cell implements Mutable {
  int _v = 0;
  int get value {
    return _v;
  }
  set value(int v) {
    _v = v;
  }
}
void main() {
  var c = Cell();
  c.value = 8;
  print(c.value);
}"#,
        ["8"]
    };

    abstract_class_concrete_method_reimplemented => {
        r#"abstract class Base {
  String prefix() {
    return 'pre';
  }
  String suffix();
}
class Impl implements Base {
  String prefix() {
    return 'pre';
  }
  String suffix() {
    return 'suf';
  }
}
void main() {
  var i = Impl();
  print(i.prefix() + i.suffix());
}"#,
        ["presuf"]
    };

    two_classes_implement_same_interface => {
        r#"abstract class Speak {
  String say();
}
class Cat implements Speak {
  String say() {
    return 'meow';
  }
}
class Dog implements Speak {
  String say() {
    return 'woof';
  }
}
void main() {
  print(Cat().say());
}"#,
        ["meow"]
    };

    interface_method_with_parameters => {
        r#"abstract class Math {
  int add(int a, int b);
}
class Calc implements Math {
  int add(int a, int b) {
    return a + b;
  }
}
void main() {
  print(Calc().add(4, 5));
}"#,
        ["9"]
    };

    implements_three_interfaces => {
        r#"abstract class A {
  int a();
}
abstract class B {
  int b();
}
abstract class C {
  int c();
}
class All implements A, B, C {
  int a() {
    return 1;
  }
  int b() {
    return 2;
  }
  int c() {
    return 3;
  }
}
void main() {
  var x = All();
  print(x.a() + x.b() + x.c());
}"#,
        ["6"]
    };

    abstract_interface_with_multiple_methods => {
        r#"abstract class Repo {
  void save(String k);
  String load(String k);
}
class MapRepo implements Repo {
  String store = '';
  void save(String k) {
    store = k;
  }
  String load(String k) {
    return store;
  }
}
void main() {
  var r = MapRepo();
  r.save('key');
  print(r.load('key'));
}"#,
        ["key"]
    };

    subclass_implements_parent_interface => {
        r#"abstract class I {
  int get n;
}
class Base implements I {
  int get n {
    return 1;
  }
}
class Sub extends Base {
  int get n {
    return 2;
  }
}
void main() {
  print(Sub().n);
}"#,
        ["2"]
    };

    interface_arrow_method_body => {
        r#"abstract class Double {
  int twice(int n);
}
class Fast implements Double {
  int twice(int n) => n * 2;
}
void main() {
  print(Fast().twice(7));
}"#,
        ["14"]
    };

    implements_after_extends_super_initializer => {
        r#"class Base {
  int x;
  Base(this.x);
}
abstract class HasY {
  int y();
}
class Pair extends Base implements HasY {
  int _y;
  Pair(int a, int b) : super(a), _y = b;
  int y() {
    return _y;
  }
}
void main() {
  var p = Pair(1, 2);
  print(p.x + p.y());
}"#,
        ["3"]
    };

    abstract_method_returning_string => {
        r#"abstract class Named {
  String name();
}
class Item implements Named {
  String name() {
    return 'item';
  }
}
void main() {
  print(Item().name());
}"#,
        ["item"]
    };

    interface_void_method_implementation => {
        r#"abstract class Action {
  void run();
}
class Job implements Action {
  int done = 0;
  void run() {
    done = 1;
  }
}
void main() {
  var j = Job();
  j.run();
  print(j.done);
}"#,
        ["1"]
    };

    two_interface_methods_same_class => {
        r#"abstract class Reader {
  int read();
}
abstract class Writer {
  int write();
}
class RW implements Reader, Writer {
  int read() {
    return 10;
  }
  int write() {
    return 20;
  }
}
void main() {
  var rw = RW();
  print(rw.read() + rw.write());
}"#,
        ["30"]
    };

    abstract_class_as_interface_no_extends => {
        r#"abstract class Fly {
  String fly();
}
class Bird implements Fly {
  String fly() {
    return 'flap';
  }
}
void main() {
  print(Bird().fly());
}"#,
        ["flap"]
    };

    implements_interface_using_instance_state => {
        r#"abstract class Counter {
  int count();
}
class Tally implements Counter {
  int _n = 0;
  void inc() {
    _n = _n + 1;
  }
  int count() {
    return _n;
  }
}
void main() {
  var t = Tally();
  t.inc();
  t.inc();
  print(t.count());
}"#,
        ["2"]
    };

    interface_with_bool_return => {
        r#"abstract class Check {
  bool ok();
}
class Pass implements Check {
  bool ok() {
    return true;
  }
}
void main() {
  print(Pass().ok());
}"#,
        ["true"]
    };

    nested_implements_with_local_helper => {
        r#"abstract class Format {
  String fmt(int n);
}
class NumFmt implements Format {
  String fmt(int n) {
    return 'n=$n';
  }
}
void main() {
  print(NumFmt().fmt(42));
}"#,
        ["n=42"]
    };

    interface_implementation_with_constructor => {
        r#"abstract class Greeter {
  String greet();
}
class Hello implements Greeter {
  String who;
  Hello(this.who);
  String greet() {
    return 'hi $who';
  }
}
void main() {
  print(Hello('Bob').greet());
}"#,
        ["hi Bob"]
    };

    multiple_implements_methods_do_not_collide => {
        r#"abstract class Left {
  int left();
}
abstract class Right {
  int right();
}
class Both implements Left, Right {
  int left() {
    return 1;
  }
  int right() {
    return 10;
  }
}
void main() {
  print(Both().left() + Both().right());
}"#,
        ["11"]
    };

    abstract_interface_only_declares_contract => {
        r#"abstract class Service {
  String ping();
}
class Live implements Service {
  String ping() {
    return 'pong';
  }
}
void main() {
  print(Live().ping());
}"#,
        ["pong"]
    };

    implements_with_override_annotation => {
        r#"abstract class Base {
  int val();
}
class Impl implements Base {
  @override
  int val() {
    return 99;
  }
}
void main() {
  print(Impl().val());
}"#,
        ["99"]
    };

    interface_method_called_via_typed_reference => {
        r#"abstract class Op {
  int run(int n);
}
class Square implements Op {
  int run(int n) {
    return n * n;
  }
}
void main() {
  Op o = Square();
  print(o.run(5));
}"#,
        ["25"]
    };

    abstract_getter_and_method_combo => {
        r#"abstract class Profile {
  String get name;
  String label();
}
class User implements Profile {
  String get name {
    return 'u1';
  }
  String label() {
    return 'user';
  }
}
void main() {
  var u = User();
  print(u.name + u.label());
}"#,
        ["u1user"]
    };

    implements_different_from_extends => {
        r#"class Base {
  int baseOnly() {
    return 1;
  }
}
abstract class Port {
  int portVal();
}
class Svc extends Base implements Port {
  int portVal() {
    return 2;
  }
}
void main() {
  var s = Svc();
  print(s.baseOnly() + s.portVal());
}"#,
        ["3"]
    };

    interface_with_static_helper_on_implementation => {
        r#"abstract class Parse {
  int parse(String s);
}
class IntParse implements Parse {
  int parse(String s) {
    return int.parse(s);
  }
}
void main() {
  print(IntParse().parse('17'));
}"#,
        ["17"]
    };
}
