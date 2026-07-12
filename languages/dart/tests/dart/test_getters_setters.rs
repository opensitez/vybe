//! Instance getters and setters: computed properties, private backing fields,
//! validation, static accessors, and override semantics.

dart_cases! {
    simple_getter_reads_backing_field => {
        r#"class Counter {
  int _count = 0;
  int get count {
    return _count;
  }
}
void main() {
  print(Counter().count);
}"#,
        ["0"]
    };

    simple_setter_writes_backing_field => {
        r#"class Counter {
  int _count = 0;
  set count(int v) {
    _count = v;
  }
  int read() {
    return _count;
  }
}
void main() {
  var c = Counter();
  c.count = 5;
  print(c.read());
}"#,
        ["5"]
    };

    getter_setter_pair_round_trip => {
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
  b.size = 10;
  print(b.size);
}"#,
        ["10"]
    };

    computed_getter_multiplies_two_fields => {
        r#"class Rect {
  int width = 3;
  int height = 4;
  int get area {
    return width * height;
  }
}
void main() {
  print(Rect().area);
}"#,
        ["12"]
    };

    computed_getter_sums_field_values => {
        r#"class Range {
  int start = 2;
  int end = 8;
  int get span {
    return end - start;
  }
}
void main() {
  print(Range().span);
}"#,
        ["6"]
    };

    arrow_body_getter_returns_expression => {
        r#"class Point {
  int x = 2;
  int y = 3;
  int get sum => x + y;
}
void main() {
  print(Point().sum);
}"#,
        ["5"]
    };

    setter_validates_non_negative_value => {
        r#"class Gauge {
  int _level = 0;
  int get level {
    return _level;
  }
  set level(int v) {
    if (v < 0) {
      v = 0;
    }
    _level = v;
  }
}
void main() {
  var g = Gauge();
  g.level = -5;
  print(g.level);
}"#,
        ["0"]
    };

    setter_clamps_to_maximum => {
        r#"class Volume {
  int _value = 0;
  int get value {
    return _value;
  }
  set value(int v) {
    if (v > 100) {
      v = 100;
    }
    _value = v;
  }
}
void main() {
  var v = Volume();
  v.value = 150;
  print(v.value);
}"#,
        ["100"]
    };

    private_backing_field_not_accessible_by_name => {
        r#"class Secret {
  int _hidden = 99;
  int reveal() {
    return _hidden;
  }
}
void main() {
  print(Secret().reveal());
}"#,
        ["99"]
    };

    getter_lazy_initializes_backing_field => {
        r#"class Lazy {
  String? _cache;
  String get label {
    _cache ??= 'ready';
    return _cache!;
  }
}
void main() {
  print(Lazy().label);
}"#,
        ["ready"]
    };

    getter_reflects_setter_mutation => {
        r#"class Label {
  String _text = 'a';
  String get text {
    return _text;
  }
  set text(String v) {
    _text = v;
  }
}
void main() {
  var l = Label();
  l.text = 'updated';
  print(l.text);
}"#,
        ["updated"]
    };

    boolean_getter_is_empty_pattern => {
        r#"class Bag {
  List<int> _items = [];
  bool get isEmpty {
    return _items.isEmpty;
  }
}
void main() {
  print(Bag().isEmpty);
}"#,
        ["true"]
    };

    boolean_getter_is_not_empty_after_add => {
        r#"class Bag {
  List<int> _items = [];
  bool get isEmpty {
    return _items.isEmpty;
  }
  void add(int v) {
    _items.add(v);
  }
}
void main() {
  var b = Bag();
  b.add(1);
  print(b.isEmpty);
}"#,
        ["false"]
    };

    static_getter_reads_static_field => {
        r#"class Config {
  static int _port = 8080;
  static int get port {
    return _port;
  }
}
void main() {
  print(Config.port);
}"#,
        ["8080"]
    };

    static_setter_updates_static_field => {
        r#"class Config {
  static int _port = 8080;
  static int get port {
    return _port;
  }
  static set port(int v) {
    _port = v;
  }
}
void main() {
  Config.port = 3000;
  print(Config.port);
}"#,
        ["3000"]
    };

    override_getter_uses_subclass_computation => {
        r#"class Animal {
  String get sound {
    return '...';
  }
}
class Dog extends Animal {
  String get sound {
    return 'woof';
  }
}
void main() {
  print(Dog().sound);
}"#,
        ["woof"]
    };

    override_setter_updates_subclass_storage => {
        r#"class Base {
  int _v = 0;
  int get v {
    return _v;
  }
  set v(int x) {
    _v = x;
  }
}
class Sub extends Base {
  set v(int x) {
    _v = x * 2;
  }
}
void main() {
  var s = Sub();
  s.v = 5;
  print(s.v);
}"#,
        ["10"]
    };

    getter_derived_from_other_getter => {
        r#"class Circle {
  double _radius = 2.0;
  double get radius {
    return _radius;
  }
  double get diameter {
    return radius * 2;
  }
}
void main() {
  print(Circle().diameter);
}"#,
        ["4.0"]
    };

    setter_triggers_side_effect_counter => {
        r#"class Tracker {
  int _writes = 0;
  int _val = 0;
  int get writes {
    return _writes;
  }
  set val(int v) {
    _writes = _writes + 1;
    _val = v;
  }
}
void main() {
  var t = Tracker();
  t.val = 1;
  t.val = 2;
  print(t.writes);
}"#,
        ["2"]
    };

    getter_only_exposes_read_access => {
        r#"class ReadOnly {
  final String _id = 'fixed';
  String get id {
    return _id;
  }
}
void main() {
  print(ReadOnly().id);
}"#,
        ["fixed"]
    };

    setter_only_with_internal_read_method => {
        r#"class WriteOnly {
  int _token = 0;
  set token(int v) {
    _token = v;
  }
  int peek() {
    return _token;
  }
}
void main() {
  var w = WriteOnly();
  w.token = 42;
  print(w.peek());
}"#,
        ["42"]
    };

    getter_string_representation_from_fields => {
        r#"class User {
  String first = 'Ada';
  String last = 'Lovelace';
  String get fullName {
    return first + ' ' + last;
  }
}
void main() {
  print(User().fullName);
}"#,
        ["Ada Lovelace"]
    };

    setter_normalizes_string_to_uppercase => {
        r#"class Code {
  String _value = '';
  String get value {
    return _value;
  }
  set value(String v) {
    _value = v.toUpperCase();
  }
}
void main() {
  var c = Code();
  c.value = 'abc';
  print(c.value);
}"#,
        ["ABC"]
    };

    getter_on_class_with_multiple_setters_via_methods => {
        r#"class Account {
  int _balance = 100;
  int get balance {
    return _balance;
  }
  void deposit(int amount) {
    _balance = _balance + amount;
  }
  void withdraw(int amount) {
    _balance = _balance - amount;
  }
}
void main() {
  var a = Account();
  a.deposit(50);
  a.withdraw(30);
  print(a.balance);
}"#,
        ["120"]
    };

    private_field_initialized_in_constructor => {
        r#"class Token {
  late String _secret;
  Token(String s) {
    _secret = s;
  }
  String get secret {
    return _secret;
  }
}
void main() {
  print(Token('key').secret);
}"#,
        ["key"]
    };

    getter_returns_bool_from_comparison => {
        r#"class Threshold {
  int level = 10;
  bool get isHigh {
    return level > 5;
  }
}
void main() {
  print(Threshold().isHigh);
}"#,
        ["true"]
    };

    setter_updates_dependent_getter => {
        r#"class Pair {
  int _a = 1;
  int _b = 2;
  int get a {
    return _a;
  }
  set a(int v) {
    _a = v;
  }
  int get total {
    return _a + _b;
  }
}
void main() {
  var p = Pair();
  p.a = 5;
  print(p.total);
}"#,
        ["7"]
    };

    static_getter_computed_from_static_data => {
        r#"class App {
  static List<String> _tags = ['dart'];
  static int get tagCount {
    return _tags.length;
  }
}
void main() {
  print(App.tagCount);
}"#,
        ["1"]
    };

    getter_with_nullable_backing_coalesces => {
        r#"class MaybeName {
  String? _name;
  String get display {
    return _name ?? 'anonymous';
  }
  set display(String v) {
    _name = v;
  }
}
void main() {
  var m = MaybeName();
  print(m.display);
  m.display = 'Zara';
  print(m.display);
}"#,
        ["anonymous", "Zara"]
    };

    cascade_setter_then_getter => {
        r#"class Widget {
  int _width = 0;
  int get width {
    return _width;
  }
  set width(int v) {
    _width = v;
  }
}
void main() {
  var w = Widget();
  w..width = 20..width = 30;
  print(w.width);
}"#,
        ["30"]
    };

    getter_exposes_length_of_private_list => {
        r#"class Collection {
  List<int> _items = [1, 2, 3];
  int get length {
    return _items.length;
  }
}
void main() {
  print(Collection().length);
}"#,
        ["3"]
    };

    setter_appends_to_private_list => {
        r#"class Collection {
  List<int> _items = [];
  set last(int v) {
    _items.add(v);
  }
  String dump() {
    return _items.join(',');
  }
}
void main() {
  var c = Collection();
  c.last = 7;
  c.last = 8;
  print(c.dump());
}"#,
        ["7,8"]
    };

    override_getter_calls_super_getter_in_expression => {
        r#"class Base {
  String get label {
    return 'base';
  }
}
class Derived extends Base {
  String get label {
    return super.label + '-ext';
  }
}
void main() {
  print(Derived().label);
}"#,
        ["base-ext"]
    };

    getter_double_value_from_int_field => {
        r#"class Scaler {
  int _factor = 3;
  int get factor {
    return _factor;
  }
  double get scaled {
    return _factor * 1.5;
  }
}
void main() {
  print(Scaler().scaled);
}"#,
        ["4.5"]
    };

    setter_and_getter_on_same_private_field_independent_instances => {
        r#"class Cell {
  int _data = 0;
  int get data {
    return _data;
  }
  set data(int v) {
    _data = v;
  }
}
void main() {
  var a = Cell();
  var b = Cell();
  a.data = 3;
  b.data = 7;
  print(a.data);
  print(b.data);
}"#,
        ["3", "7"]
    };
}
