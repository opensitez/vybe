//! Dart constructor forms: default, named, redirecting, factory, const,
//! and initializer lists.

dart_cases! {
    default_constructor_creates_usable_instance => {
        r#"class Widget {}
void main() {
  var w = Widget();
  print(w != null);
}"#,
        ["true"]
    };

    generative_positional_constructor_sets_fields => {
        r#"class Pair {
  int x;
  int y;
  Pair(this.x, this.y);
}
void main() {
  var p = Pair(3, 4);
  print(p.x + p.y);
}"#,
        ["7"]
    };

    this_shorthand_assigns_multiple_fields => {
        r#"class RGB {
  int r;
  int g;
  int b;
  RGB(this.r, this.g, this.b);
}
void main() {
  var c = RGB(1, 2, 3);
  print(c.r + c.g + c.b);
}"#,
        ["6"]
    };

    named_constructor_sets_alternate_initial_state => {
        r#"class Point {
  int x;
  int y;
  Point(this.x, this.y);
  Point.origin() : x = 0, y = 0;
}
void main() {
  var p = Point.origin();
  print(p.x);
}"#,
        ["0"]
    };

    named_constructor_with_body_computes_values => {
        r#"class Size {
  int w;
  int h;
  Size(this.w, this.h);
  Size.square(int side) : w = side, h = side;
}
void main() {
  var s = Size.square(5);
  print(s.w + s.h);
}"#,
        ["10"]
    };

    redirecting_constructor_delegates_to_primary => {
        r#"class Point {
  int x;
  int y;
  Point(this.x, this.y);
  Point.zero() : this(0, 0);
}
void main() {
  var p = Point.zero();
  print(p.y);
}"#,
        ["0"]
    };

    redirecting_named_to_other_named => {
        r#"class Vector {
  int x;
  int y;
  Vector(this.x, this.y);
  Vector.unit() : this(1, 1);
  Vector.zero() : this.unit();
}
void main() {
  var v = Vector.zero();
  print(v.x);
}"#,
        ["1"]
    };

    factory_constructor_returns_new_instance => {
        r#"class Box {
  int value;
  Box(this.value);
  factory Box.empty() {
    return Box(0);
  }
}
void main() {
  print(Box.empty().value);
}"#,
        ["0"]
    };

    factory_constructor_can_return_cached_instance => {
        r#"class Cache {
  static Cache? _one;
  int id;
  Cache._(this.id);
  factory Cache() {
    _one ??= Cache._(1);
    return _one!;
  }
}
void main() {
  var a = Cache();
  var b = Cache();
  print(a == b);
}"#,
        ["true"]
    };

    factory_redirects_to_private_named_constructor => {
        r#"class Token {
  String text;
  Token._(this.text);
  factory Token.fromText(String t) {
    return Token._(t);
  }
}
void main() {
  print(Token.fromText('ok').text);
}"#,
        ["ok"]
    };

    const_constructor_allows_const_instance => {
        r#"class Imm {
  final int x;
  const Imm(this.x);
}
void main() {
  const v = Imm(7);
  print(v.x);
}"#,
        ["7"]
    };

    const_constructor_multiple_fields => {
        r#"class Pair {
  final int a;
  final int b;
  const Pair(this.a, this.b);
}
void main() {
  const p = Pair(2, 3);
  print(p.a + p.b);
}"#,
        ["5"]
    };

    initializer_list_assigns_field_before_body => {
        r#"class Account {
  int balance;
  Account(int start) : balance = start {
    balance = balance + 1;
  }
}
void main() {
  print(Account(10).balance);
}"#,
        ["11"]
    };

    initializer_list_with_super_call => {
        r#"class Base {
  int n;
  Base(this.n);
}
class Sub extends Base {
  Sub(int v) : super(v);
}
void main() {
  print(Sub(12).n);
}"#,
        ["12"]
    };

    initializer_list_multiple_field_assignments => {
        r#"class Line {
  int x1;
  int y1;
  int x2;
  int y2;
  Line(int a, int b, int c, int d)
      : x1 = a, y1 = b, x2 = c, y2 = d;
}
void main() {
  var l = Line(0, 0, 3, 4);
  print(l.x2 + l.y2);
}"#,
        ["7"]
    };

    generative_constructor_with_assert_in_initializer => {
        r#"class Positive {
  int n;
  Positive(this.n) : assert(n > 0);
}
void main() {
  print(Positive(5).n);
}"#,
        ["5"]
    };

    named_constructor_on_subclass_calls_super => {
        r#"class Base {
  int n;
  Base(this.n);
}
class Sub extends Base {
  Sub(int v) : super(v);
  Sub.zero() : super(0);
}
void main() {
  print(Sub.zero().n);
}"#,
        ["0"]
    };

    factory_with_logic_selects_implementation => {
        r#"class Shape {
  int sides;
  Shape(this.sides);
  factory Shape.triangle() {
    return Shape(3);
  }
  factory Shape.square() {
    return Shape(4);
  }
}
void main() {
  print(Shape.triangle().sides);
}"#,
        ["3"]
    };

    factory_named_alternate_entry_point => {
        r#"class Id {
  int value;
  Id(this.value);
  factory Id.zero() {
    return Id(0);
  }
}
void main() {
  print(Id.zero().value);
}"#,
        ["0"]
    };

    const_named_constructor => {
        r#"class Origin {
  final int x;
  final int y;
  const Origin(this.x, this.y);
  const Origin.zero() : x = 0, y = 0;
}
void main() {
  const o = Origin.zero();
  print(o.y);
}"#,
        ["0"]
    };

    constructor_optional_positional_default => {
        r#"class Msg {
  String text;
  Msg([this.text = 'default']);
}
void main() {
  print(Msg().text);
}"#,
        ["default"]
    };

    constructor_named_optional_defaults => {
        r#"class Config {
  int port;
  Config({this.port = 8080});
}
void main() {
  print(Config().port);
}"#,
        ["8080"]
    };

    initializer_list_computes_from_parameters => {
        r#"class Span {
  int start;
  int end;
  Span(int len) : start = 0, end = len;
}
void main() {
  print(Span(5).end);
}"#,
        ["5"]
    };

    super_initializer_before_subclass_field_init => {
        r#"class Base {
  int a;
  Base(this.a);
}
class Sub extends Base {
  int b;
  Sub(int x, int y) : super(x), b = y;
}
void main() {
  var s = Sub(2, 3);
  print(s.a + s.b);
}"#,
        ["5"]
    };

    factory_returns_subclass_instance => {
        r#"class Animal {
  String kind;
  Animal(this.kind);
  factory Animal.dog() {
    return Dog('dog');
  }
}
class Dog extends Animal {
  Dog(String k) : super(k);
}
void main() {
  print(Animal.dog().kind);
}"#,
        ["dog"]
    };

    private_generative_with_public_factory => {
        r#"class Secret {
  int code;
  Secret._(this.code);
  factory Secret.open(int c) {
    return Secret._(c);
  }
}
void main() {
  print(Secret.open(99).code);
}"#,
        ["99"]
    };

    redirecting_to_primary_with_different_args => {
        r#"class Range {
  int lo;
  int hi;
  Range(this.lo, this.hi);
  Range.single(int n) : this(n, n);
}
void main() {
  var r = Range.single(4);
  print(r.hi);
}"#,
        ["4"]
    };

    const_constructor_used_in_static_const => {
        r#"class Vec {
  final int x;
  final int y;
  const Vec(this.x, this.y);
  static const zero = Vec(0, 0);
}
void main() {
  print(Vec.zero.x);
}"#,
        ["0"]
    };

    constructor_body_runs_after_initializers => {
        r#"class Log {
  int step;
  Log(int s) : step = s {
    step = step + 10;
  }
}
void main() {
  print(Log(1).step);
}"#,
        ["11"]
    };

    named_constructor_with_super_initializer => {
        r#"class Base {
  int n;
  Base(this.n);
}
class Child extends Base {
  Child(int v) : super(v);
  Child.empty() : super(0);
}
void main() {
  print(Child.empty().n);
}"#,
        ["0"]
    };

    factory_constructor_with_multiple_creation_paths => {
        r#"class Num {
  int v;
  Num(this.v);
  factory Num.parse(String s) {
    if (s == 'zero') {
      return Num(0);
    }
    return Num(int.parse(s));
  }
}
void main() {
  print(Num.parse('zero').v);
}"#,
        ["0"]
    };

    initializer_list_super_then_fields => {
        r#"class A {
  int x;
  A(this.x);
}
class B extends A {
  int y;
  B(int a, int b) : super(a), y = b;
}
void main() {
  var b = B(1, 2);
  print(b.x + b.y);
}"#,
        ["3"]
    };

    generative_constructor_no_fields => {
        r#"class Marker {
  bool ready = true;
}
void main() {
  print(Marker().ready);
}"#,
        ["true"]
    };

    redirecting_chain_three_constructors => {
        r#"class N {
  int v;
  N(this.v);
  N.one() : this(1);
  N.copy() : this.one();
}
void main() {
  print(N.copy().v);
}"#,
        ["1"]
    };

    factory_with_static_state => {
        r#"class Seq {
  static int _n = 0;
  int id;
  Seq(this.id);
  factory Seq.next() {
    _n = _n + 1;
    return Seq(_n);
  }
}
void main() {
  print(Seq.next().id);
}"#,
        ["1"]
    };

    const_constructor_field_access => {
        r#"class ConstBox {
  final int w;
  final int h;
  const ConstBox(this.w, this.h);
  int get area {
    return w * h;
  }
}
void main() {
  const b = ConstBox(2, 5);
  print(b.area);
}"#,
        ["10"]
    };

    initializer_assigns_from_expression => {
        r#"class Square {
  int side;
  int area;
  Square(int s) : side = s, area = s * s;
}
void main() {
  print(Square(6).area);
}"#,
        ["36"]
    };

    subclass_generative_calls_implicit_super => {
        r#"class Base {
  int n = 5;
}
class Sub extends Base {}
void main() {
  print(Sub().n);
}"#,
        ["5"]
    };

    factory_named_vs_generative_named_distinction => {
        r#"class Point {
  int x;
  int y;
  Point(this.x, this.y);
  Point.fromXY(int a, int b) : x = a, y = b;
  factory Point.middle() {
    return Point(50, 50);
  }
}
void main() {
  print(Point.middle().x);
}"#,
        ["50"]
    };

    generative_constructor_assigns_in_body_after_params => {
        r#"class Slot {
  int value;
  Slot(int seed) {
    value = seed + 1;
  }
}
void main() {
  print(Slot(4).value);
}"#,
        ["5"]
    };
}
