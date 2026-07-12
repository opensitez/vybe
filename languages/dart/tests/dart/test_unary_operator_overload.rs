//! Unary operator overloading: operator - (negation) and operator ~ (bitwise NOT).

dart_cases! {
    unary_minus_negates_positive_value => {
        r#"class Signed {
  int v;
  Signed(this.v);
  Signed operator -() {
    return Signed(-v);
  }
}
void main() {
  print((-Signed(5)).v);
}"#,
        ["-5"]
    };

    unary_minus_negates_negative_value => {
        r#"class Signed {
  int v;
  Signed(this.v);
  Signed operator -() {
    return Signed(-v);
  }
}
void main() {
  print((-Signed(-3)).v);
}"#,
        ["3"]
    };

    unary_minus_on_zero => {
        r#"class Num {
  int n;
  Num(this.n);
  Num operator -() {
    return Num(-n);
  }
}
void main() {
  print((-Num(0)).n);
}"#,
        ["0"]
    };

    unary_minus_double_negation => {
        r#"class Val {
  int x;
  Val(this.x);
  Val operator -() {
    return Val(-x);
  }
}
void main() {
  print((-(-Val(7))).x);
}"#,
        ["7"]
    };

    unary_minus_preserves_type => {
        r#"class Counter {
  int count;
  Counter(this.count);
  Counter operator -() {
    return Counter(-count);
  }
}
void main() {
  var c = Counter(10);
  var n = -c;
  print(n.count);
}"#,
        ["-10"]
    };

    unary_minus_used_in_expression => {
        r#"class Point {
  int x;
  Point(this.x);
  Point operator -() {
    return Point(-x);
  }
}
void main() {
  print((-Point(4)).x + 1);
}"#,
        ["-3"]
    };

    unary_minus_in_equality_check => {
        r#"class Score {
  int pts;
  Score(this.pts);
  Score operator -() {
    return Score(-pts);
  }
}
void main() {
  print((-Score(8)).pts == -8);
}"#,
        ["true"]
    };

    unary_minus_large_value => {
        r#"class Big {
  int v;
  Big(this.v);
  Big operator -() {
    return Big(-v);
  }
}
void main() {
  print((-Big(1000)).v);
}"#,
        ["-1000"]
    };

    unary_minus_chained_three_times => {
        r#"class Flip {
  int n;
  Flip(this.n);
  Flip operator -() {
    return Flip(-n);
  }
}
void main() {
  print((-(-(-Flip(2)))).n);
}"#,
        ["-2"]
    };

    unary_minus_with_field_mutation_after => {
        r#"class Box {
  int v;
  Box(this.v);
  Box operator -() {
    return Box(-v);
  }
}
void main() {
  var b = Box(6);
  var neg = -b;
  b.v = 1;
  print(neg.v);
}"#,
        ["-6"]
    };

    unary_bitwise_not_flips_bits => {
        r#"class Mask {
  int bits;
  Mask(this.bits);
  Mask operator ~() {
    return Mask(~bits);
  }
}
void main() {
  print((~Mask(0)).bits);
}"#,
        ["-1"]
    };

    unary_bitwise_not_on_255 => {
        r#"class Byte {
  int v;
  Byte(this.v);
  Byte operator ~() {
    return Byte(~v);
  }
}
void main() {
  print((~Byte(255)).v);
}"#,
        ["-256"]
    };

    unary_bitwise_not_double_application => {
        r#"class Bits {
  int n;
  Bits(this.n);
  Bits operator ~() {
    return Bits(~n);
  }
}
void main() {
  print((~(~Bits(42))).n);
}"#,
        ["42"]
    };

    unary_bitwise_not_on_zero => {
        r#"class Word {
  int w;
  Word(this.w);
  Word operator ~() {
    return Word(~w);
  }
}
void main() {
  print((~Word(0)).w);
}"#,
        ["-1"]
    };

    unary_bitwise_not_on_negative => {
        r#"class Reg {
  int r;
  Reg(this.r);
  Reg operator ~() {
    return Reg(~r);
  }
}
void main() {
  print((~Reg(-1)).r);
}"#,
        ["0"]
    };

    unary_bitwise_not_in_addition => {
        r#"class Flag {
  int f;
  Flag(this.f);
  Flag operator ~() {
    return Flag(~f);
  }
}
void main() {
  print((~Flag(0)).f + 1);
}"#,
        ["0"]
    };

    unary_bitwise_not_preserves_wrapper => {
        r#"class Slot {
  int s;
  Slot(this.s);
  Slot operator ~() {
    return Slot(~s);
  }
}
void main() {
  var a = Slot(5);
  var b = ~a;
  print(b.s);
}"#,
        ["-6"]
    };

    unary_bitwise_not_on_one => {
        r#"class Bit {
  int b;
  Bit(this.b);
  Bit operator ~() {
    return Bit(~b);
  }
}
void main() {
  print((~Bit(1)).b);
}"#,
        ["-2"]
    };

    unary_minus_and_bitwise_not_combined => {
        r#"class Dual {
  int v;
  Dual(this.v);
  Dual operator -() {
    return Dual(-v);
  }
  Dual operator ~() {
    return Dual(~v);
  }
}
void main() {
  print((-Dual(3)).v);
  print((~Dual(3)).v);
}"#,
        ["-3", "-4"]
    };

    unary_minus_on_custom_negative_class => {
        r#"class Negative {
  int value;
  Negative(this.value);
  Negative operator -() {
    return Negative(-value);
  }
  bool isNegative() {
    return value < 0;
  }
}
void main() {
  var n = Negative(-9);
  var pos = -n;
  print(pos.value);
  print(pos.isNegative());
}"#,
        ["9", "false"]
    };

    unary_minus_returns_new_instance => {
        r#"class Pair {
  int a;
  Pair(this.a);
  Pair operator -() {
    return Pair(-a);
  }
}
void main() {
  var p = Pair(2);
  var q = -p;
  print(p.a);
  print(q.a);
}"#,
        ["2", "-2"]
    };

    unary_bitwise_not_returns_new_instance => {
        r#"class Cell {
  int c;
  Cell(this.c);
  Cell operator ~() {
    return Cell(~c);
  }
}
void main() {
  var a = Cell(10);
  var b = ~a;
  print(a.c);
  print(b.c);
}"#,
        ["10", "-11"]
    };

    unary_minus_in_list_literal => {
        r#"class N {
  int v;
  N(this.v);
  N operator -() {
    return N(-v);
  }
}
void main() {
  var items = [-N(1), -N(2)];
  print(items[0].v + items[1].v);
}"#,
        ["-3"]
    };

    unary_bitwise_not_in_conditional => {
        r#"class Gate {
  int g;
  Gate(this.g);
  Gate operator ~() {
    return Gate(~g);
  }
}
void main() {
  var x = ~Gate(0);
  print(x.g == -1);
}"#,
        ["true"]
    };

    unary_minus_on_field_access_chain => {
        r#"class Outer {
  Inner inner;
  Outer(this.inner);
}
class Inner {
  int v;
  Inner(this.v);
  Inner operator -() {
    return Inner(-v);
  }
}
void main() {
  print((-Outer(Inner(5)).inner).v);
}"#,
        ["-5"]
    };

    unary_minus_with_compare_to_zero => {
        r#"class Magnitude {
  int m;
  Magnitude(this.m);
  Magnitude operator -() {
    return Magnitude(-m);
  }
}
void main() {
  print((-Magnitude(0)).m == 0);
}"#,
        ["true"]
    };

    unary_bitwise_not_hex_pattern => {
        r#"class Hex {
  int h;
  Hex(this.h);
  Hex operator ~() {
    return Hex(~h);
  }
}
void main() {
  print((~Hex(0x0F)).h);
}"#,
        ["-16"]
    };

    unary_minus_in_print_interpolation => {
        r#"class Tag {
  int id;
  Tag(this.id);
  Tag operator -() {
    return Tag(-id);
  }
}
void main() {
  var t = -Tag(7);
  print(t.id);
}"#,
        ["-7"]
    };

    unary_bitwise_not_small_positive => {
        r#"class Tiny {
  int t;
  Tiny(this.t);
  Tiny operator ~() {
    return Tiny(~t);
  }
}
void main() {
  print((~Tiny(2)).t);
}"#,
        ["-3"]
    };

    unary_minus_from_method_return => {
        r#"class Source {
  int n;
  Source(this.n);
  Source operator -() {
    return Source(-n);
  }
  Source make() {
    return Source(4);
  }
}
void main() {
  print((-Source(0).make()).n);
}"#,
        ["-4"]
    };

    unary_bitwise_not_from_variable => {
        r#"class Store {
  int s;
  Store(this.s);
  Store operator ~() {
    return Store(~s);
  }
}
void main() {
  var base = Store(8);
  var flipped = ~base;
  print(flipped.s);
}"#,
        ["-9"]
    };

    unary_minus_twice_equals_original => {
        r#"class Rev {
  int r;
  Rev(this.r);
  Rev operator -() {
    return Rev(-r);
  }
}
void main() {
  var orig = Rev(11);
  var back = -(-orig);
  print(back.r);
}"#,
        ["11"]
    };

    unary_bitwise_not_twice_restores => {
        r#"class Toggle {
  int t;
  Toggle(this.t);
  Toggle operator ~() {
    return Toggle(~t);
  }
}
void main() {
  print((~(~Toggle(99))).t);
}"#,
        ["99"]
    };

    unary_minus_negative_class_name_semantics => {
        r#"class Balance {
  int amount;
  Balance(this.amount);
  Balance operator -() {
    return Balance(-amount);
  }
}
void main() {
  var debt = Balance(-100);
  var credit = -debt;
  print(credit.amount);
}"#,
        ["100"]
    };

    unary_bitwise_not_on_max_small_int => {
        r#"class Lim {
  int l;
  Lim(this.l);
  Lim operator ~() {
    return Lim(~l);
  }
}
void main() {
  print((~Lim(127)).l);
}"#,
        ["-128"]
    };

    unary_minus_in_equality_with_literal => {
        r#"class Unit {
  int u;
  Unit(this.u);
  Unit operator -() {
    return Unit(-u);
  }
}
void main() {
  print((-Unit(5)).u == -5);
}"#,
        ["true"]
    };

    unary_bitwise_not_result_is_negative => {
        r#"class Pos {
  int p;
  Pos(this.p);
  Pos operator ~() {
    return Pos(~p);
  }
}
void main() {
  print((~Pos(0)).p < 0);
}"#,
        ["true"]
    };

    unary_minus_on_subclass_field => {
        r#"class Base {
  int v;
  Base(this.v);
}
class Derived extends Base {
  Derived(int v) : super(v);
  Derived operator -() {
    return Derived(-v);
  }
}
void main() {
  print((-Derived(3)).v);
}"#,
        ["-3"]
    };

    unary_bitwise_not_used_in_if => {
        r#"class Switch {
  int s;
  Switch(this.s);
  Switch operator ~() {
    return Switch(~s);
  }
}
void main() {
  var r = ~Switch(0);
  if (r.s == -1) {
    print('ok');
  }
}"#,
        ["ok"]
    };

    unary_minus_and_plus_binary_distinct => {
        r#"class AddNeg {
  int n;
  AddNeg(this.n);
  AddNeg operator -() {
    return AddNeg(-n);
  }
  AddNeg operator +(AddNeg o) {
    return AddNeg(n + o.n);
  }
}
void main() {
  var a = AddNeg(5);
  print((-a).n);
  print((a + AddNeg(1)).n);
}"#,
        ["-5", "6"]
    };
}
