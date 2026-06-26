//! Field initialization: declaration defaults, initializer lists,
//! this.field assignments, and evaluation order.

dart_cases! {
    declaration_initializer_int_literal => {
        r#"class Cell {
  int value = 42;
}
void main() {
  print(Cell().value);
}"#,
        ["42"]
    };

    declaration_initializer_string_literal => {
        r#"class Tag {
  String label = 'alpha';
}
void main() {
  print(Tag().label);
}"#,
        ["alpha"]
    };

    declaration_initializer_bool_true => {
        r#"class Flag {
  bool on = true;
}
void main() {
  print(Flag().on);
}"#,
        ["true"]
    };

    declaration_initializer_list_literal => {
        r#"class Bucket {
  List<int> items = [1, 2, 3];
}
void main() {
  print(Bucket().items.length);
}"#,
        ["3"]
    };

    declaration_initializer_map_literal => {
        r#"class Registry {
  Map<String, int> scores = {'a': 10};
}
void main() {
  print(Registry().scores['a']);
}"#,
        ["10"]
    };

    declaration_initializer_arithmetic_expression => {
        r#"class Scale {
  int factor = 2 + 3;
}
void main() {
  print(Scale().factor);
}"#,
        ["5"]
    };

    declaration_initializer_nullable_defaults_null => {
        r#"class Maybe {
  String? name;
}
void main() {
  print(Maybe().name);
}"#,
        ["null"]
    };

    initializer_list_overrides_declaration_default => {
        r#"class Override {
  int n = 1;
  Override(int v) : n = v;
}
void main() {
  print(Override(9).n);
}"#,
        ["9"]
    };

    initializer_list_assigns_uninitialized_field => {
        r#"class Point {
  int x;
  int y;
  Point(int a, int b) : x = a, y = b;
}
void main() {
  var p = Point(3, 4);
  print(p.x + p.y);
}"#,
        ["7"]
    };

    this_field_shorthand_in_initializer_list => {
        r#"class Pair {
  int a;
  int b;
  Pair(int x, int y) : a = x, b = y;
}
void main() {
  print(Pair(2, 5).b);
}"#,
        ["5"]
    };

    initializer_list_this_field_from_parameter => {
        r#"class Holder {
  int value;
  Holder(this.value);
}
void main() {
  print(Holder(17).value);
}"#,
        ["17"]
    };

    initializer_list_field_references_prior_field => {
        r#"class Rect {
  int width;
  int height;
  int area;
  Rect(int w, int h) : width = w, height = h, area = w * h;
}
void main() {
  print(Rect(4, 5).area);
}"#,
        ["20"]
    };

    initializer_list_second_field_uses_first => {
        r#"class Span {
  int start = 0;
  int end;
  Span(int len) : end = start + len;
}
void main() {
  print(Span(7).end);
}"#,
        ["7"]
    };

    initializer_list_computes_from_constructor_params => {
        r#"class Circle {
  int radius;
  int diameter;
  Circle(int r) : radius = r, diameter = r * 2;
}
void main() {
  print(Circle(6).diameter);
}"#,
        ["12"]
    };

    body_runs_after_initializer_list_assignment => {
        r#"class Step {
  int count;
  Step(int seed) : count = seed {
    count = count + 100;
  }
}
void main() {
  print(Step(3).count);
}"#,
        ["103"]
    };

    declaration_then_initializer_list_then_body_order => {
        r#"class Trace {
  int stage = 1;
  Trace(int bump) : stage = stage + bump {
    stage = stage + 1000;
  }
}
void main() {
  print(Trace(2).stage);
}"#,
        ["1003"]
    };

    super_initializer_runs_before_subclass_field_init => {
        r#"class Base {
  int baseVal;
  Base(this.baseVal);
}
class Sub extends Base {
  int subVal;
  Sub(int b, int s) : super(b), subVal = s;
}
void main() {
  var s = Sub(10, 20);
  print(s.baseVal + s.subVal);
}"#,
        ["30"]
    };

    super_call_before_this_field_in_subclass_list => {
        r#"class Parent {
  int p;
  Parent(this.p);
}
class Child extends Parent {
  int c;
  Child(int x, int y) : super(x), c = y + 1;
}
void main() {
  print(Child(5, 10).c);
}"#,
        ["11"]
    };

    named_constructor_initializer_list_sets_fields => {
        r#"class Vector {
  int x;
  int y;
  Vector(this.x, this.y);
  Vector.zero() : x = 0, y = 0;
}
void main() {
  print(Vector.zero().y);
}"#,
        ["0"]
    };

    named_constructor_initializer_differs_from_primary => {
        r#"class Size {
  int w;
  int h;
  Size(this.w, this.h);
  Size.square(int side) : w = side, h = side;
}
void main() {
  var s = Size.square(6);
  print(s.w + s.h);
}"#,
        ["12"]
    };

    redirecting_constructor_preserves_initializer_chain => {
        r#"class Id {
  int value;
  Id(this.value);
  Id.zero() : this(0);
}
void main() {
  print(Id.zero().value);
}"#,
        ["0"]
    };

    final_field_set_only_in_initializer_list => {
        r#"class Token {
  final String code;
  Token(String c) : code = c;
}
void main() {
  print(Token('abc').code);
}"#,
        ["abc"]
    };

    final_field_via_this_shorthand => {
        r#"class Code {
  final int n;
  Code(this.n);
}
void main() {
  print(Code(88).n);
}"#,
        ["88"]
    };

    initializer_list_with_assert_before_body => {
        r#"class Positive {
  int n;
  Positive(int v) : n = v, assert(v > 0);
}
void main() {
  print(Positive(4).n);
}"#,
        ["4"]
    };

    initializer_list_multiple_asserts_with_fields => {
        r#"class Range {
  int lo;
  int hi;
  Range(int a, int b) : lo = a, hi = b, assert(a <= b);
}
void main() {
  print(Range(2, 8).hi);
}"#,
        ["8"]
    };

    field_init_uses_static_constant => {
        r#"class Config {
  static const defaultPort = 8080;
  int port = defaultPort;
}
void main() {
  print(Config().port);
}"#,
        ["8080"]
    };

    instance_field_init_does_not_share_mutable_list => {
        r#"class Tray {
  List<int> slots = [];
}
void main() {
  var a = Tray();
  var b = Tray();
  a.slots.add(1);
  print(b.slots.length);
}"#,
        ["0"]
    };

    subclass_inherits_parent_declaration_initializer => {
        r#"class Base {
  int n = 5;
}
class Sub extends Base {}
void main() {
  print(Sub().n);
}"#,
        ["5"]
    };

    subclass_field_init_after_super_default_constructor => {
        r#"class Base {
  int a = 1;
}
class Sub extends Base {
  int b = 2;
}
void main() {
  var s = Sub();
  print(s.a + s.b);
}"#,
        ["3"]
    };

    subclass_initializer_list_after_implicit_super => {
        r#"class Base {
  int x = 10;
}
class Sub extends Base {
  int y;
  Sub(int v) : y = v;
}
void main() {
  print(Sub(3).x + Sub(3).y);
}"#,
        ["13"]
    };

    initializer_assigns_field_from_superclass_method => {
        r#"class Base {
  int baseDouble(int n) {
    return n * 2;
  }
}
class Sub extends Base {
  int stored;
  Sub(int seed) : stored = seed + 1;
}
void main() {
  print(Sub(4).stored);
}"#,
        ["5"]
    };

    three_field_initializer_list_left_to_right => {
        r#"class Triple {
  int a;
  int b;
  int c;
  Triple(int x) : a = x, b = x + 1, c = x + 2;
}
void main() {
  var t = Triple(10);
  print(t.a + t.b + t.c);
}"#,
        ["33"]
    };

    initializer_list_string_interpolation_from_param => {
        r#"class Greet {
  String msg;
  Greet(String name) : msg = 'hi $name';
}
void main() {
  print(Greet('Ann').msg);
}"#,
        ["hi Ann"]
    };

    field_default_plus_initializer_list_override => {
        r#"class Meter {
  int reading = 1;
  Meter.reset() : reading = 0;
}
void main() {
  print(Meter.reset().reading);
}"#,
        ["0"]
    };

    constructor_body_assigns_after_all_initializers => {
        r#"class Ledger {
  int balance = 0;
  Ledger(int deposit) : balance = deposit {
    balance = balance - 1;
  }
}
void main() {
  print(Ledger(50).balance);
}"#,
        ["49"]
    };

    initializer_list_with_super_and_multiple_fields => {
        r#"class A {
  int u;
  A(this.u);
}
class B extends A {
  int v;
  int w;
  B(int a, int b, int c) : super(a), v = b, w = c;
}
void main() {
  var b = B(1, 2, 3);
  print(b.u + b.v + b.w);
}"#,
        ["6"]
    };

    late_field_not_initialized_at_declaration => {
        r#"class Lazy {
  late int value;
  Lazy(int v) : value = v;
}
void main() {
  print(Lazy(9).value);
}"#,
        ["9"]
    };

    field_initializer_negation_expression => {
        r#"class Sign {
  int flipped = -5;
}
void main() {
  print(Sign().flipped);
}"#,
        ["-5"]
    };

    initializer_list_division_truncates => {
        r#"class Ratio {
  int whole;
  Ratio(int a, int b) : whole = a ~/ b;
}
void main() {
  print(Ratio(7, 2).whole);
}"#,
        ["3"]
    };

    field_init_bool_from_comparison => {
        r#"class Check {
  bool ok = 3 > 2;
}
void main() {
  print(Check().ok);
}"#,
        ["true"]
    };

    initializer_list_sets_both_coords_from_single_param => {
        r#"class Point1D {
  int x;
  int y;
  Point1D(int v) : x = v, y = v;
}
void main() {
  print(Point1D(8).x + Point1D(8).y);
}"#,
        ["16"]
    };

    declaration_init_double_literal => {
        r#"class Measure {
  double pi = 3.14;
}
void main() {
  print(Measure().pi > 3.0);
}"#,
        ["true"]
    };

    initializer_list_chained_named_constructor => {
        r#"class Box {
  int size;
  Box(this.size);
  Box.small() : size = 1;
  Box.large() : size = 100;
}
void main() {
  print(Box.small().size + Box.large().size);
}"#,
        ["101"]
    };

    field_init_empty_string_default => {
        r#"class Buffer {
  String text = '';
}
void main() {
  print(Buffer().text.length);
}"#,
        ["0"]
    };

    initializer_list_this_multiple_fields => {
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

    super_named_constructor_with_subclass_field_init => {
        r#"class Base {
  int n;
  Base(this.n);
  Base.zero() : n = 0;
}
class Sub extends Base {
  int extra;
  Sub.zero(int e) : super.zero(), extra = e;
}
void main() {
  var s = Sub.zero(5);
  print(s.n + s.extra);
}"#,
        ["5"]
    };

    field_init_set_literal => {
        r#"class Unique {
  Set<int> ids = {1, 2};
}
void main() {
  print(Unique().ids.length);
}"#,
        ["2"]
    };

    initializer_list_field_from_conditional_expression => {
        r#"class Pick {
  int chosen;
  Pick(bool useA, int a, int b) : chosen = useA ? a : b;
}
void main() {
  print(Pick(false, 1, 9).chosen);
}"#,
        ["9"]
    };

    declaration_and_initializer_both_contribute_to_order => {
        r#"class Order {
  int first = 1;
  int second;
  Order(int bump) : second = first + bump;
}
void main() {
  print(Order(4).second);
}"#,
        ["5"]
    };

    subclass_overrides_declaration_with_own_field_default => {
        r#"class Base {
  int level = 1;
}
class Sub extends Base {
  int level = 10;
}
void main() {
  print(Sub().level);
}"#,
        ["10"]
    };
}
