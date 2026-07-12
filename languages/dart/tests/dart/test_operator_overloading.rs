//! Custom operator +, operator ==, and operator [] overloading.

dart_cases! {
    operator_plus_adds_two_int_fields => {
        r#"class Vec2 {
  int x;
  int y;
  Vec2(this.x, this.y);
  Vec2 operator +(Vec2 other) {
    return Vec2(x + other.x, y + other.y);
  }
}
void main() {
  var a = Vec2(1, 2);
  var b = Vec2(3, 4);
  var c = a + b;
  print(c.x + c.y);
}"#,
        ["10"]
    };

    operator_plus_with_zero_identity => {
        r#"class Num {
  int v;
  Num(this.v);
  Num operator +(Num other) {
    return Num(v + other.v);
  }
}
void main() {
  var n = Num(5) + Num(0);
  print(n.v);
}"#,
        ["5"]
    };

    operator_plus_chained_three_times => {
        r#"class Adder {
  int n;
  Adder(this.n);
  Adder operator +(Adder other) {
    return Adder(n + other.n);
  }
}
void main() {
  var r = Adder(1) + Adder(2) + Adder(3);
  print(r.n);
}"#,
        ["6"]
    };

    operator_plus_string_concat_fields => {
        r#"class Tag {
  String a;
  String b;
  Tag(this.a, this.b);
  Tag operator +(Tag other) {
    return Tag(a + other.a, b + other.b);
  }
}
void main() {
  var t = Tag('x', '1') + Tag('y', '2');
  print(t.a + t.b);
}"#,
        ["xy12"]
    };

    operator_equals_same_values_true => {
        r#"class Point {
  int x;
  int y;
  Point(this.x, this.y);
  bool operator ==(Object other) {
    if (other is Point) {
      return x == other.x && y == other.y;
    }
    return false;
  }
}
void main() {
  print(Point(1, 2) == Point(1, 2));
}"#,
        ["true"]
    };

    operator_equals_different_values_false => {
        r#"class Point {
  int x;
  int y;
  Point(this.x, this.y);
  bool operator ==(Object other) {
    if (other is Point) {
      return x == other.x && y == other.y;
    }
    return false;
  }
}
void main() {
  print(Point(1, 2) == Point(3, 4));
}"#,
        ["false"]
    };

    operator_equals_non_matching_type_false => {
        r#"class Box {
  int v;
  Box(this.v);
  bool operator ==(Object other) {
    return other is Box && v == other.v;
  }
}
void main() {
  print(Box(1) == 1);
}"#,
        ["false"]
    };

    operator_equals_same_instance_true => {
        r#"class Item {
  int id;
  Item(this.id);
  bool operator ==(Object other) {
    return identical(this, other);
  }
}
void main() {
  var a = Item(1);
  print(a == a);
}"#,
        ["true"]
    };

    operator_index_read_first_slot => {
        r#"class Row {
  List<int> cells;
  Row(this.cells);
  int operator [](int i) {
    return cells[i];
  }
}
void main() {
  print(Row([10, 20, 30])[0]);
}"#,
        ["10"]
    };

    operator_index_read_last_slot => {
        r#"class Row {
  List<int> cells;
  Row(this.cells);
  int operator [](int i) {
    return cells[i];
  }
}
void main() {
  print(Row([10, 20, 30])[2]);
}"#,
        ["30"]
    };

    operator_index_assign_mutates_backing_list => {
        r#"class Grid {
  List<int> data;
  Grid(this.data);
  int operator [](int i) {
    return data[i];
  }
  void operator []=(int i, int v) {
    data[i] = v;
  }
}
void main() {
  var g = Grid([1, 2, 3]);
  g[1] = 99;
  print(g[1]);
}"#,
        ["99"]
    };

    operator_plus_used_in_print_expression => {
        r#"class W {
  int n;
  W(this.n);
  W operator +(W o) {
    return W(n + o.n);
  }
}
void main() {
  print((W(2) + W(3)).n);
}"#,
        ["5"]
    };

    operator_equals_single_field => {
        r#"class Id {
  int value;
  Id(this.value);
  bool operator ==(Object other) {
    return other is Id && value == other.value;
  }
}
void main() {
  print(Id(7) == Id(7));
}"#,
        ["true"]
    };

    operator_index_on_string_list => {
        r#"class Words {
  List<String> items;
  Words(this.items);
  String operator [](int i) {
    return items[i];
  }
}
void main() {
  print(Words(['a', 'b'])[1]);
}"#,
        ["b"]
    };

    operator_plus_returns_new_instance => {
        r#"class Pair {
  int a;
  int b;
  Pair(this.a, this.b);
  Pair operator +(Pair o) {
    return Pair(a + o.a, b + o.b);
  }
}
void main() {
  var p = Pair(1, 1) + Pair(2, 3);
  print(p.a);
}"#,
        ["3"]
    };

    operator_equals_reflexive_on_equal_fields => {
        r#"class RGB {
  int r;
  int g;
  int b;
  RGB(this.r, this.g, this.b);
  bool operator ==(Object other) {
    if (other is! RGB) return false;
    return r == other.r && g == other.g && b == other.b;
  }
}
void main() {
  var c = RGB(1, 2, 3);
  print(c == RGB(1, 2, 3));
}"#,
        ["true"]
    };

    operator_index_middle_element => {
        r#"class Tape {
  List<int> vals;
  Tape(this.vals);
  int operator [](int i) => vals[i];
}
void main() {
  print(Tape([5, 6, 7])[1]);
}"#,
        ["6"]
    };

    operator_plus_with_negative_values => {
        r#"class Signed {
  int v;
  Signed(this.v);
  Signed operator +(Signed o) {
    return Signed(v + o.v);
  }
}
void main() {
  print((Signed(-2) + Signed(5)).v);
}"#,
        ["3"]
    };

    operator_equals_after_mutation_false => {
        r#"class Score {
  int pts;
  Score(this.pts);
  bool operator ==(Object other) {
    return other is Score && pts == other.pts;
  }
}
void main() {
  var a = Score(1);
  var b = Score(1);
  b.pts = 2;
  print(a == b);
}"#,
        ["false"]
    };

    operator_index_assign_then_read => {
        r#"class Buffer {
  List<int> buf;
  Buffer(this.buf);
  int operator [](int i) => buf[i];
  void operator []=(int i, int v) {
    buf[i] = v;
  }
}
void main() {
  var b = Buffer([0, 0]);
  b[0] = 42;
  print(b[0]);
}"#,
        ["42"]
    };

    operator_plus_commutative_values => {
        r#"class Val {
  int n;
  Val(this.n);
  Val operator +(Val o) => Val(n + o.n);
}
void main() {
  print((Val(10) + Val(1)).n);
}"#,
        ["11"]
    };

    operator_equals_null_object_false => {
        r#"class Node {
  int id;
  Node(this.id);
  bool operator ==(Object other) {
    if (other == null) return false;
    return other is Node && id == other.id;
  }
}
void main() {
  print(Node(1) == null);
}"#,
        ["false"]
    };

    operator_index_zero_based_second => {
        r#"class Slots {
  List<String> s;
  Slots(this.s);
  String operator [](int i) {
    return s[i];
  }
}
void main() {
  print(Slots(['first', 'second'])[1]);
}"#,
        ["second"]
    };

    operator_plus_large_sum => {
        r#"class Big {
  int v;
  Big(this.v);
  Big operator +(Big o) {
    return Big(v + o.v);
  }
}
void main() {
  print((Big(1000) + Big(234)).v);
}"#,
        ["1234"]
    };

    operator_equals_symmetric_pair => {
        r#"class Token {
  String key;
  Token(this.key);
  bool operator ==(Object other) {
    return other is Token && key == other.key;
  }
}
void main() {
  var a = Token('k');
  var b = Token('k');
  print(a == b);
}"#,
        ["true"]
    };

    operator_index_sum_of_two_reads => {
        r#"class Duo {
  List<int> pair;
  Duo(this.pair);
  int operator [](int i) {
    return pair[i];
  }
}
void main() {
  var d = Duo([3, 7]);
  print(d[0] + d[1]);
}"#,
        ["10"]
    };

    operator_plus_preserves_second_operand_field => {
        r#"class Mix {
  int x;
  int y;
  Mix(this.x, this.y);
  Mix operator +(Mix o) {
    return Mix(x + o.x, y);
  }
}
void main() {
  print((Mix(1, 9) + Mix(2, 0)).y);
}"#,
        ["9"]
    };

    operator_equals_int_wrapped_value => {
        r#"class Wrap {
  int inner;
  Wrap(this.inner);
  bool operator ==(Object other) {
    return other is Wrap && inner == other.inner;
  }
}
void main() {
  print(Wrap(0) == Wrap(0));
}"#,
        ["true"]
    };

    operator_index_assign_multiple_cells => {
        r#"class Matrix {
  List<int> flat;
  Matrix(this.flat);
  int operator [](int i) => flat[i];
  void operator []=(int i, int v) {
    flat[i] = v;
  }
}
void main() {
  var m = Matrix([0, 0, 0]);
  m[0] = 1;
  m[2] = 3;
  print(m[0] + m[2]);
}"#,
        ["4"]
    };

    operator_plus_double_application => {
        r#"class Inc {
  int n;
  Inc(this.n);
  Inc operator +(Inc o) {
    return Inc(n + o.n);
  }
}
void main() {
  var base = Inc(1);
  var step = Inc(4);
  print((base + step + step).n);
}"#,
        ["9"]
    };
}
