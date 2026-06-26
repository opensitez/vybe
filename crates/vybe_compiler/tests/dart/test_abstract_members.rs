//! Abstract classes: abstract methods, getters, setters; concrete subclasses
//! must implement; abstract types used polymorphically.

dart_cases! {
    abstract_method_implemented_by_subclass => {
        r#"abstract class Shape {
  double area();
}
class Square extends Shape {
  int side;
  Square(this.side);
  double area() {
    return side * side;
  }
}
void main() {
  print(Square(4).area());
}"#,
        ["16.0"]
    };

    abstract_getter_implemented_by_subclass => {
        r#"abstract class Named {
  String get name;
}
class User extends Named {
  String get name => 'Alice';
}
void main() {
  print(User().name);
}"#,
        ["Alice"]
    };

    abstract_setter_implemented_by_subclass => {
        r#"abstract class Mutable {
  set value(int v);
  int get value;
}
class Counter extends Mutable {
  int _v = 0;
  set value(int v) {
    _v = v;
  }
  int get value => _v;
}
void main() {
  var c = Counter();
  c.value = 9;
  print(c.value);
}"#,
        ["9"]
    };

    abstract_method_and_getter_both_implemented => {
        r#"abstract class Entity {
  String get id;
  String describe();
}
class Product extends Entity {
  String get id => 'p1';
  String describe() {
    return 'product';
  }
}
void main() {
  print(Product().describe());
}"#,
        ["product"]
    };

    abstract_class_with_concrete_method_inherited => {
        r#"abstract class Base {
  int baseVal() {
    return 10;
  }
  int extra();
}
class Sub extends Base {
  int extra() {
    return 5;
  }
}
void main() {
  print(Sub().baseVal() + Sub().extra());
}"#,
        ["15"]
    };

    abstract_concrete_method_used_without_override => {
        r#"abstract class Logger {
  void log(String msg) {
    print('log:$msg');
  }
  void flush();
}
class FileLogger extends Logger {
  void flush() {
    print('flushed');
  }
}
void main() {
  FileLogger().log('x');
}"#,
        ["log:x"]
    };

    abstract_field_with_concrete_getter => {
        r#"abstract class HasCount {
  int get count;
  void increment() {
    print('inc');
  }
}
class LiveCount extends HasCount {
  int _n = 0;
  int get count => _n;
}
void main() {
  print(LiveCount().count);
}"#,
        ["0"]
    };

    abstract_multiple_methods_all_implemented => {
        r#"abstract class Calc {
  int add(int a, int b);
  int sub(int a, int b);
}
class BasicCalc extends Calc {
  int add(int a, int b) {
    return a + b;
  }
  int sub(int a, int b) {
    return a - b;
  }
}
void main() {
  print(BasicCalc().add(7, 3));
}"#,
        ["10"]
    };

    abstract_polymorphic_call_through_supertype => {
        r#"abstract class Greeter {
  String greet();
}
class EnGreeter extends Greeter {
  String greet() {
    return 'hello';
  }
}
void main() {
  Greeter g = EnGreeter();
  print(g.greet());
}"#,
        ["hello"]
    };

    abstract_two_subclasses_different_impl => {
        r#"abstract class Op {
  int run(int n);
}
class Double extends Op {
  int run(int n) {
    return n * 2;
  }
}
class Triple extends Op {
  int run(int n) {
    return n * 3;
  }
}
void main() {
  print(Double().run(4) + Triple().run(4));
}"#,
        ["20"]
    };

    abstract_method_returns_bool => {
        r#"abstract class Check {
  bool ok();
}
class AlwaysYes extends Check {
  bool ok() {
    return true;
  }
}
void main() {
  print(AlwaysYes().ok());
}"#,
        ["true"]
    };

    abstract_method_returns_list => {
        r#"abstract class Provider {
  List<int> provide();
}
class Range extends Provider {
  List<int> provide() {
    return [1, 2, 3];
  }
}
void main() {
  print(Range().provide().length);
}"#,
        ["3"]
    };

    abstract_getter_computed_from_field => {
        r#"abstract class Sized {
  int get width;
  int get height;
  int area() {
    return width * height;
  }
}
class Rect extends Sized {
  int w;
  int h;
  Rect(this.w, this.h);
  int get width => w;
  int get height => h;
}
void main() {
  print(Rect(3, 4).area());
}"#,
        ["12"]
    };

    abstract_setter_triggers_side_effect => {
        r#"abstract class Store {
  set token(String t);
  String get token;
}
class MemStore extends Store {
  String _t = '';
  set token(String t) {
    _t = t;
  }
  String get token => _t;
}
void main() {
  var s = MemStore();
  s.token = 'abc';
  print(s.token.length);
}"#,
        ["3"]
    };

    abstract_class_constructor_in_subclass => {
        r#"abstract class Vehicle {
  String make;
  Vehicle(this.make);
  void drive();
}
class Car extends Vehicle {
  Car(String m) : super(m);
  void drive() {
    print(make);
  }
}
void main() {
  Car('Vybe').drive();
}"#,
        ["Vybe"]
    };

    abstract_three_level_inheritance => {
        r#"abstract class A {
  int a();
}
abstract class B extends A {
  int b();
}
class C extends B {
  int a() {
    return 1;
  }
  int b() {
    return 2;
  }
}
void main() {
  print(C().a() + C().b());
}"#,
        ["3"]
    };

    abstract_method_with_parameters => {
        r#"abstract class Formatter {
  String fmt(String input, int width);
}
class Pad extends Formatter {
  String fmt(String input, int width) {
    return input + ':$width';
  }
}
void main() {
  print(Pad().fmt('x', 5));
}"#,
        ["x:5"]
    };

    abstract_void_method_implementation => {
        r#"abstract class Runner {
  void run();
}
class Sprint extends Runner {
  void run() {
    print('fast');
  }
}
void main() {
  Sprint().run();
}"#,
        ["fast"]
    };

    abstract_getter_only_subclass => {
        r#"abstract class Config {
  String get appName;
}
class DevConfig extends Config {
  String get appName => 'dev';
}
void main() {
  print(DevConfig().appName);
}"#,
        ["dev"]
    };

    abstract_method_string_concat => {
        r#"abstract class Builder {
  String build();
}
class GreetBuilder extends Builder {
  String name;
  GreetBuilder(this.name);
  String build() {
    return 'hi $name';
  }
}
void main() {
  print(GreetBuilder('Bob').build());
}"#,
        ["hi Bob"]
    };

    abstract_list_in_concrete_method => {
        r#"abstract class Collector {
  List<String> items = [];
  void collect(String s) {
    items.add(s);
  }
  int total();
}
class SimpleCollector extends Collector {
  int total() {
    return items.length;
  }
}
void main() {
  var c = SimpleCollector();
  c.collect('a');
  c.collect('b');
  print(c.total());
}"#,
        ["2"]
    };

    abstract_static_method_on_abstract_class => {
        r#"abstract class MathUtil {
  static int doubleIt(int n) {
    return n * 2;
  }
  int compute();
}
class Doubler extends MathUtil {
  int compute() {
    return MathUtil.doubleIt(5);
  }
}
void main() {
  print(Doubler().compute());
}"#,
        ["10"]
    };

    abstract_instance_used_in_list => {
        r#"abstract class Node {
  int val();
}
class Leaf extends Node {
  int v;
  Leaf(this.v);
  int val() {
    return v;
  }
}
void main() {
  List<Node> nodes = [Leaf(1), Leaf(2)];
  print(nodes[0].val() + nodes[1].val());
}"#,
        ["3"]
    };

    abstract_method_called_from_concrete_method => {
        r#"abstract class Tax {
  double rate();
  double apply(double amount) {
    return amount * rate();
  }
}
class SalesTax extends Tax {
  double rate() {
    return 0.1;
  }
}
void main() {
  print(SalesTax().apply(100.0));
}"#,
        ["10.0"]
    };

    abstract_getter_setter_pair => {
        r#"abstract class Temperature {
  double get celsius;
  set celsius(double v);
}
class Room extends Temperature {
  double _c = 20.0;
  double get celsius => _c;
  set celsius(double v) {
    _c = v;
  }
}
void main() {
  var r = Room();
  r.celsius = 25.0;
  print(r.celsius);
}"#,
        ["25.0"]
    };

    abstract_subclass_overrides_concrete => {
        r#"abstract class Base {
  String tag() {
    return 'base';
  }
  String label();
}
class Sub extends Base {
  String label() {
    return super.tag() + '-sub';
  }
}
void main() {
  print(Sub().label());
}"#,
        ["base-sub"]
    };

    abstract_method_null_return => {
        r#"abstract class Finder {
  String? find(int id);
}
class NullFinder extends Finder {
  String? find(int id) {
    return null;
  }
}
void main() {
  print(NullFinder().find(1) == null);
}"#,
        ["true"]
    };

    abstract_method_finds_value => {
        r#"abstract class Finder {
  String? find(int id);
}
class MapFinder extends Finder {
  String? find(int id) {
    if (id == 1) {
      return 'found';
    }
    return null;
  }
}
void main() {
  print(MapFinder().find(1));
}"#,
        ["found"]
    };

    abstract_class_with_final_field_in_sub => {
        r#"abstract class IdEntity {
  String get id;
}
class Record extends IdEntity {
  final String _id;
  Record(this._id);
  String get id => _id;
}
void main() {
  print(Record('r42').id);
}"#,
        ["r42"]
    };

    abstract_multiple_getters => {
        r#"abstract class Point2D {
  int get x;
  int get y;
}
class Vec extends Point2D {
  int _x;
  int _y;
  Vec(this._x, this._y);
  int get x => _x;
  int get y => _y;
}
void main() {
  print(Vec(3, 7).x + Vec(3, 7).y);
}"#,
        ["10"]
    };

    abstract_method_in_switch => {
        r#"abstract class Status {
  String code();
}
class Ok extends Status {
  String code() {
    return 'ok';
  }
}
class Err extends Status {
  String code() {
    return 'err';
  }
}
void main() {
  Status s = Ok();
  print(s.code());
}"#,
        ["ok"]
    };

    abstract_implement_two_abstract_methods => {
        r#"abstract class RW {
  String read();
  void write(String s);
}
class Buffer extends RW {
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
  b.write('data');
  print(b.read());
}"#,
        ["data"]
    };

    abstract_concrete_default_used_by_sub => {
        r#"abstract class Parser {
  String parse(String input) {
    return input.trim();
  }
  String format(String s);
}
class SimpleParser extends Parser {
  String format(String s) {
    return parse(s).toUpperCase();
  }
}
void main() {
  print(SimpleParser().format('  hi  '));
}"#,
        ["HI"]
    };

    abstract_method_recursive_fib => {
        r#"abstract class Seq {
  int at(int n);
}
class Fib extends Seq {
  int at(int n) {
    if (n <= 1) {
      return n;
    }
    return at(n - 1) + at(n - 2);
  }
}
void main() {
  print(Fib().at(6));
}"#,
        ["8"]
    };

    abstract_getter_lazy_init => {
        r#"abstract class Lazy {
  String get data;
}
class LazyImpl extends Lazy {
  String? _cache;
  String get data {
    _cache ??= 'loaded';
    return _cache!;
  }
}
void main() {
  print(LazyImpl().data);
}"#,
        ["loaded"]
    };

    abstract_method_with_optional_param => {
        r#"abstract class Greeter {
  String hi([String name = 'world']);
}
class Polite extends Greeter {
  String hi([String name = 'world']) {
    return 'hello $name';
  }
}
void main() {
  print(Polite().hi());
}"#,
        ["hello world"]
    };

    abstract_method_named_params => {
        r#"abstract class Connect {
  bool open({required String host, int port = 80});
}
class Tcp extends Connect {
  bool open({required String host, int port = 80}) {
    print('$host:$port');
    return true;
  }
}
void main() {
  print(Tcp().open(host: 'localhost', port: 8080));
}"#,
        ["localhost:8080", "true"]
    };

    abstract_subclass_adds_method => {
        r#"abstract class Animal {
  String speak();
}
class Dog extends Animal {
  String speak() {
    return 'woof';
  }
  String fetch() {
    return 'ball';
  }
}
void main() {
  print(Dog().speak() + Dog().fetch());
}"#,
        ["woofball"]
    };

    abstract_field_in_subclass_not_abstract => {
        r#"abstract class Widget {
  void render();
}
class Button extends Widget {
  String label = 'click';
  void render() {
    print(label);
  }
}
void main() {
  Button().render();
}"#,
        ["click"]
    };

    abstract_method_returns_map => {
        r#"abstract class Mapper {
  Map<String, int> map();
}
class ScoreMap extends Mapper {
  Map<String, int> map() {
    return {'a': 1, 'b': 2};
  }
}
void main() {
  print(ScoreMap().map()['b']);
}"#,
        ["2"]
    };

    abstract_hierarchy_diamond_methods => {
        r#"abstract class Top {
  int top();
}
abstract class Left extends Top {
  int left();
}
class Bottom extends Left {
  int top() {
    return 1;
  }
  int left() {
    return 10;
  }
}
void main() {
  print(Bottom().top() + Bottom().left());
}"#,
        ["11"]
    };

    abstract_method_bool_negation => {
        r#"abstract class Gate {
  bool open();
}
class Locked extends Gate {
  bool open() {
    return false;
  }
}
void main() {
  print(!Locked().open());
}"#,
        ["true"]
    };

    abstract_getter_from_constructor_param => {
        r#"abstract class Named {
  String get name;
}
class Person extends Named {
  final String _name;
  Person(this._name);
  String get name => _name;
}
void main() {
  print(Person('Eve').name);
}"#,
        ["Eve"]
    };

    abstract_method_divide => {
        r#"abstract class Divider {
  double div(double a, double b);
}
class Halve extends Divider {
  double div(double a, double b) {
    return a / b;
  }
}
void main() {
  print(Halve().div(10.0, 4.0));
}"#,
        ["2.5"]
    };

    abstract_concrete_calls_abstract => {
        r#"abstract class Template {
  String step1() {
    return 'a';
  }
  String step2();
  String run() {
    return step1() + step2();
  }
}
class Impl extends Template {
  String step2() {
    return 'b';
  }
}
void main() {
  print(Impl().run());
}"#,
        ["ab"]
    };

    abstract_multiple_instances => {
        r#"abstract class Id {
  int get id;
}
class A extends Id {
  int get id => 1;
}
class B extends Id {
  int get id => 2;
}
void main() {
  print(A().id + B().id);
}"#,
        ["3"]
    };

    abstract_method_empty_string => {
        r#"abstract class Empty {
  String value();
}
class Blank extends Empty {
  String value() {
    return '';
  }
}
void main() {
  print(Blank().value().length);
}"#,
        ["0"]
    };

    abstract_setter_validation => {
        r#"abstract class Validated {
  set age(int v);
  int get age;
}
class Person extends Validated {
  int _age = 0;
  set age(int v) {
    if (v >= 0) {
      _age = v;
    }
  }
  int get age => _age;
}
void main() {
  var p = Person();
  p.age = 30;
  print(p.age);
}"#,
        ["30"]
    };

    abstract_method_max_of_two => {
        r#"abstract class Max {
  int max(int a, int b);
}
class IntMax extends Max {
  int max(int a, int b) {
    return a > b ? a : b;
  }
}
void main() {
  print(IntMax().max(12, 7));
}"#,
        ["12"]
    };

    abstract_getter_bool_flag => {
        r#"abstract class Feature {
  bool get enabled;
}
class Beta extends Feature {
  bool get enabled => true;
}
void main() {
  print(Beta().enabled);
}"#,
        ["true"]
    };
}
