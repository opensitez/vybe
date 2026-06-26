//! Deep const semantics: const constructors, const collection literals,
//! compile-time arithmetic, identical canonicalized const objects, and static const fields.

dart_cases! {
    const_int_addition_at_compile_time => {
        r#"void main() {
  const sum = 10 + 20;
  print(sum);
}"#,
        ["30"]
    };

    const_int_subtraction_chain => {
        r#"void main() {
  const diff = 100 - 30 - 10;
  print(diff);
}"#,
        ["60"]
    };

    const_int_multiplication => {
        r#"void main() {
  const product = 6 * 7;
  print(product);
}"#,
        ["42"]
    };

    const_int_division_truncates => {
        r#"void main() {
  const quotient = 17 ~/ 5;
  print(quotient);
}"#,
        ["3"]
    };

    const_int_modulo => {
        r#"void main() {
  const remainder = 17 % 5;
  print(remainder);
}"#,
        ["2"]
    };

    const_bitwise_and => {
        r#"void main() {
  const masked = 0xFF & 0x0F;
  print(masked);
}"#,
        ["15"]
    };

    const_bitwise_or => {
        r#"void main() {
  const combined = 0x10 | 0x01;
  print(combined);
}"#,
        ["17"]
    };

    const_bitwise_xor => {
        r#"void main() {
  const toggled = 0xFF ^ 0x0F;
  print(toggled);
}"#,
        ["240"]
    };

    const_left_shift => {
        r#"void main() {
  const shifted = 1 << 4;
  print(shifted);
}"#,
        ["16"]
    };

    const_right_shift => {
        r#"void main() {
  const shifted = 32 >> 2;
  print(shifted);
}"#,
        ["8"]
    };

    const_negative_int_literal => {
        r#"void main() {
  const n = -42;
  print(n);
}"#,
        ["-42"]
    };

    const_hex_int_literal => {
        r#"void main() {
  const n = 0x2A;
  print(n);
}"#,
        ["42"]
    };

    const_double_addition => {
        r#"void main() {
  const sum = 1.5 + 2.5;
  print(sum);
}"#,
        ["4.0"]
    };

    const_double_multiplication => {
        r#"void main() {
  const area = 3.0 * 4.0;
  print(area);
}"#,
        ["12.0"]
    };

    const_bool_and => {
        r#"void main() {
  const flag = true && true;
  print(flag);
}"#,
        ["true"]
    };

    const_bool_or => {
        r#"void main() {
  const flag = false || true;
  print(flag);
}"#,
        ["true"]
    };

    const_bool_not => {
        r#"void main() {
  const flag = !false;
  print(flag);
}"#,
        ["true"]
    };

    const_string_concatenation => {
        r#"void main() {
  const greeting = 'Hello' + ' ' + 'Dart';
  print(greeting);
}"#,
        ["Hello Dart"]
    };

    const_string_adjacent_literals => {
        r#"void main() {
  const word = 'hel' 'lo';
  print(word);
}"#,
        ["hello"]
    };

    const_from_other_const_variable => {
        r#"void main() {
  const base = 5;
  const doubled = base * 2;
  print(doubled);
}"#,
        ["10"]
    };

    const_comparison_in_expression => {
        r#"void main() {
  const pick = 3 > 2 ? 10 : 20;
  print(pick);
}"#,
        ["10"]
    };

    const_list_literal_length => {
        r#"void main() {
  const items = [1, 2, 3, 4];
  print(items.length);
}"#,
        ["4"]
    };

    const_list_literal_first_element => {
        r#"void main() {
  const items = [10, 20, 30];
  print(items[0]);
}"#,
        ["10"]
    };

    const_list_literal_last_element => {
        r#"void main() {
  const items = [10, 20, 30];
  print(items[2]);
}"#,
        ["30"]
    };

    const_empty_list_length => {
        r#"void main() {
  const empty = <int>[];
  print(empty.length);
}"#,
        ["0"]
    };

    const_typed_list_literal => {
        r#"void main() {
  const nums = <int>[7, 8, 9];
  print(nums.join(','));
}"#,
        ["7,8,9"]
    };

    const_map_literal_lookup => {
        r#"void main() {
  const m = {'a': 1, 'b': 2};
  print(m['a']);
}"#,
        ["1"]
    };

    const_map_literal_length => {
        r#"void main() {
  const m = {'x': 10, 'y': 20, 'z': 30};
  print(m.length);
}"#,
        ["3"]
    };

    const_empty_map_length => {
        r#"void main() {
  const m = <String, int>{};
  print(m.length);
}"#,
        ["0"]
    };

    const_typed_map_literal => {
        r#"void main() {
  const m = <String, int>{'one': 1, 'two': 2};
  print(m['two']);
}"#,
        ["2"]
    };

    const_set_literal_length => {
        r#"void main() {
  const s = {1, 2, 3};
  print(s.length);
}"#,
        ["3"]
    };

    const_set_literal_contains => {
        r#"void main() {
  const s = {10, 20, 30};
  print(s.contains(20));
}"#,
        ["true"]
    };

    const_empty_set_length => {
        r#"void main() {
  const s = <int>{};
  print(s.length);
}"#,
        ["0"]
    };

    const_typed_set_literal => {
        r#"void main() {
  const s = <String>{'a', 'b'};
  print(s.contains('a'));
}"#,
        ["true"]
    };

    const_constructor_simple_fields => {
        r#"class Point {
  final int x;
  final int y;
  const Point(this.x, this.y);
}
void main() {
  const p = Point(3, 4);
  print(p.x + p.y);
}"#,
        ["7"]
    };

    const_constructor_named_fields => {
        r#"class Size {
  final int width;
  final int height;
  const Size({required this.width, required this.height});
}
void main() {
  const s = Size(width: 5, height: 6);
  print(s.width * s.height);
}"#,
        ["30"]
    };

    const_constructor_with_initializer_list => {
        r#"class Square {
  final int side;
  final int area;
  const Square(int s) : side = s, area = s * s;
}
void main() {
  const sq = Square(4);
  print(sq.area);
}"#,
        ["16"]
    };

    const_constructor_redirecting => {
        r#"class Counter {
  final int value;
  const Counter(this.value);
  const Counter.zero() : value = 0;
}
void main() {
  const c = Counter.zero();
  print(c.value);
}"#,
        ["0"]
    };

    const_constructor_nested_objects => {
        r#"class Inner {
  final int n;
  const Inner(this.n);
}
class Outer {
  final Inner inner;
  const Outer(this.inner);
}
void main() {
  const o = Outer(Inner(9));
  print(o.inner.n);
}"#,
        ["9"]
    };

    static_const_field_access => {
        r#"class Config {
  static const int timeout = 30;
  static const String host = 'localhost';
}
void main() {
  print(Config.timeout);
  print(Config.host);
}"#,
        ["30", "localhost"]
    };

    static_const_used_in_const_expression => {
        r#"class Limits {
  static const int max = 100;
}
void main() {
  const half = Limits.max ~/ 2;
  print(half);
}"#,
        ["50"]
    };

    const_class_field_initialized_at_declaration => {
        r#"class Defaults {
  static const String version = '2.0';
  static const int build = 42;
}
void main() {
  print(Defaults.version);
  print(Defaults.build);
}"#,
        ["2.0", "42"]
    };

    identical_two_const_int_literals => {
        r#"void main() {
  print(identical(42, 42));
}"#,
        ["true"]
    };

    identical_two_const_string_literals => {
        r#"void main() {
  print(identical('dart', 'dart'));
}"#,
        ["true"]
    };

    identical_two_const_list_literals => {
        r#"void main() {
  print(identical(const [1, 2], const [1, 2]));
}"#,
        ["true"]
    };

    identical_two_const_map_literals => {
        r#"void main() {
  print(identical(const {'k': 1}, const {'k': 1}));
}"#,
        ["true"]
    };

    identical_two_const_set_literals => {
        r#"void main() {
  print(identical(const {1, 2}, const {1, 2}));
}"#,
        ["true"]
    };

    identical_const_list_via_alias => {
        r#"void main() {
  const a = [1, 2, 3];
  const b = a;
  print(identical(a, b));
}"#,
        ["true"]
    };

    identical_const_constructor_instances => {
        r#"class Token {
  final int id;
  const Token(this.id);
}
void main() {
  print(identical(const Token(1), const Token(1)));
}"#,
        ["true"]
    };

    identical_empty_const_lists => {
        r#"void main() {
  print(identical(const <int>[], const <int>[]));
}"#,
        ["true"]
    };

    const_list_of_const_objects => {
        r#"class Cell {
  final int v;
  const Cell(this.v);
}
void main() {
  const row = [Cell(1), Cell(2), Cell(3)];
  print(row[1].v);
}"#,
        ["2"]
    };

    const_map_with_const_values => {
        r#"class Code {
  final int n;
  const Code(this.n);
}
void main() {
  const codes = {'a': Code(1), 'b': Code(2)};
  print(codes['b']!.n);
}"#,
        ["2"]
    };

    const_enum_member_is_canonical => {
        r#"enum Status { ok, fail }
void main() {
  const s = Status.ok;
  print(identical(s, Status.ok));
}"#,
        ["true"]
    };

    const_in_switch_case_label => {
        r#"const int kCode = 2;
void main() {
  switch (kCode) {
    case 2:
      print('matched');
      break;
    default:
      print('other');
  }
}"#,
        ["matched"]
    };

    const_duration_seconds => {
        r#"void main() {
  const d = Duration(seconds: 5);
  print(d.inSeconds);
}"#,
        ["5"]
    };

    const_duration_from_milliseconds => {
        r#"void main() {
  const d = Duration(milliseconds: 1500);
  print(d.inSeconds);
}"#,
        ["1"]
    };

    const_arithmetic_deep_chain => {
        r#"void main() {
  const result = (2 + 3) * 4 - 6 ~/ 2;
  print(result);
}"#,
        ["17"]
    };

    const_subclass_const_constructor => {
        r#"class Base {
  final int n;
  const Base(this.n);
}
class Derived extends Base {
  const Derived(int v) : super(v);
}
void main() {
  const d = Derived(11);
  print(d.n);
}"#,
        ["11"]
    };

    const_final_field_only_const_constructor => {
        r#"class Immutable {
  final int value;
  const Immutable(this.value);
}
void main() {
  const obj = Immutable(77);
  print(obj.value);
}"#,
        ["77"]
    };

    const_list_index_in_const_context => {
        r#"void main() {
  const items = [5, 10, 15];
  const third = items[2];
  print(third);
}"#,
        ["15"]
    };

    const_map_key_in_const_context => {
        r#"void main() {
  const m = {'pi': 3, 'e': 2};
  const val = m['pi'];
  print(val);
}"#,
        ["3"]
    };
}
