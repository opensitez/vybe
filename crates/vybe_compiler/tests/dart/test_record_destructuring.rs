//! Record destructuring: var (a,b)=record, named/nested destructure, swap via records, function return destructure.

dart_cases! {
    var_positional_pair_destructure => {
        r#"void main() {
  var (a, b) = (10, 20);
  print(a);
  print(b);
}"#,
        ["10", "20"]
    };

    var_positional_triple_destructure => {
        r#"void main() {
  var (x, y, z) = (1, 2, 3);
  print(x + y + z);
}"#,
        ["6"]
    };

    var_named_two_field_destructure => {
        r#"void main() {
  var (name: n, age: a) = (name: 'Ada', age: 33);
  print(n);
  print(a);
}"#,
        ["Ada", "33"]
    };

    var_mixed_positional_named_destructure => {
        r#"void main() {
  var (code, message: msg) = (404, message: 'Not Found');
  print(code);
  print(msg);
}"#,
        ["404", "Not Found"]
    };

    swap_two_ints_via_record_assignment => {
        r#"void main() {
  var a = 1;
  var b = 9;
  (a, b) = (b, a);
  print(a);
  print(b);
}"#,
        ["9", "1"]
    };

    swap_three_values_via_rotating_record => {
        r#"void main() {
  var a = 1;
  var b = 2;
  var c = 3;
  (a, b, c) = (c, a, b);
  print(a);
  print(b);
  print(c);
}"#,
        ["3", "1", "2"]
    };

    function_return_positional_destructured => {
        r#"(int, int) pair() => (7, 8);
void main() {
  var (x, y) = pair();
  print(x);
  print(y);
}"#,
        ["7", "8"]
    };

    function_return_named_destructured => {
        r#"({String city, int pop}) meta() => (city: 'Oslo', pop: 700000);
void main() {
  var (city: c, pop: p) = meta();
  print(c);
  print(p);
}"#,
        ["Oslo", "700000"]
    };

    function_return_mixed_destructured => {
        r#"(int, {String unit}) measure() => (42, unit: 'px');
void main() {
  var (val, unit: u) = measure();
  print(val);
  print(u);
}"#,
        ["42", "px"]
    };

    nested_record_outer_inner_destructure => {
        r#"void main() {
  var ((a, b), tag: t) = ((1, 2), tag: 'ok');
  print(a);
  print(b);
  print(t);
}"#,
        ["1", "2", "ok"]
    };

    nested_record_destructure_then_sum_inner => {
        r#"void main() {
  var ((x, y), (z, w)) = ((1, 2), (3, 4));
  print(x + y + z + w);
}"#,
        ["10"]
    };

    destructure_from_list_of_records_in_for => {
        r#"void main() {
  var pts = [(1, 2), (3, 4), (5, 6)];
  var sum = 0;
  for (var (x, y) in pts) {
    sum += x + y;
  }
  print(sum);
}"#,
        ["21"]
    };

    destructure_named_record_in_for_in => {
        r#"void main() {
  var users = [(name: 'Ann', score: 10), (name: 'Bob', score: 20)];
  var total = 0;
  for (var (name: _, score: s) in users) {
    total += s;
  }
  print(total);
}"#,
        ["30"]
    };

    destructure_record_from_map_value => {
        r#"void main() {
  var data = {'home': (10, 20), 'away': (30, 40)};
  var (x, y) = data['home']!;
  print(x);
  print(y);
}"#,
        ["10", "20"]
    };

    destructure_after_spread_record_literal => {
        r#"void main() {
  var base = (a: 1, b: 2);
  var (a: x, b: y, c: z) = (a: base.a, b: base.b, c: 3);
  print(x);
  print(y);
  print(z);
}"#,
        ["1", "2", "3"]
    };

    double_swap_restores_original_values => {
        r#"void main() {
  var p = 5;
  var q = 15;
  (p, q) = (q, p);
  (p, q) = (q, p);
  print(p);
  print(q);
}"#,
        ["5", "15"]
    };

    destructure_single_field_positional_record => {
        r#"void main() {
  var (only,) = (99,);
  print(only);
}"#,
        ["99"]
    };

    destructure_single_named_field_record => {
        r#"void main() {
  var (value: v) = (value: 77);
  print(v);
}"#,
        ["77"]
    };

    destructure_four_positional_fields => {
        r#"void main() {
  var (a, b, c, d) = (1, 2, 3, 4);
  print(a * b + c * d);
}"#,
        ["14"]
    };

    destructure_rgb_named_triple => {
        r#"void main() {
  var (r: red, g: green, b: blue) = (r: 10, g: 20, b: 30);
  print(red + green + blue);
}"#,
        ["60"]
    };

    sync_return_destructure_from_bounds_helper => {
        r#"({int min, int max}) bounds(List<int> xs) {
  return (min: xs.first, max: xs.last);
}
void main() {
  var (min: lo, max: hi) = bounds([3, 9, 1, 7]);
  print(lo);
  print(hi);
}"#,
        ["3", "7"]
    };

    destructure_in_variable_declaration_chain => {
        r#"void main() {
  var (a, b) = (1, 2);
  var (c, d) = (b, a);
  print(c);
  print(d);
}"#,
        ["2", "1"]
    };

    destructure_record_with_string_and_int => {
        r#"void main() {
  var (label, count: n) = ('items', count: 5);
  print(label);
  print(n);
}"#,
        ["items", "5"]
    };

    destructure_bool_and_int_named_fields => {
        r#"void main() {
  var (ok: flag, code: c) = (ok: true, code: 200);
  print(flag);
  print(c);
}"#,
        ["true", "200"]
    };

    destructure_from_function_with_record_param => {
        r#"int sumPair((int, int) p) {
  var (a, b) = p;
  return a + b;
}
void main() {
  print(sumPair((4, 6)));
}"#,
        ["10"]
    };

    destructure_returned_record_in_expression => {
        r#"(int, int) twice(int n) => (n, n * 2);
void main() {
  var (a, b) = twice(5);
  print(a + b);
}"#,
        ["15"]
    };

    destructure_mixed_after_reassignment => {
        r#"void main() {
  var rec = (id: 1, tag: 'a');
  var (id: i, tag: t) = rec;
  rec = (id: 2, tag: 'b');
  print(i);
  print(t);
}"#,
        ["1", "a"]
    };

    triple_nested_positional_destructure => {
        r#"void main() {
  var (((a))) = (((7)));
  print(a);
}"#,
        ["7"]
    };

    destructure_record_equality_after_bind => {
        r#"void main() {
  var src = (x: 1, y: 2);
  var (x: a, y: b) = src;
  print(a == src.x);
  print(b == src.y);
}"#,
        ["true", "true"]
    };

    destructure_from_conditional_record_pick => {
        r#"void main() {
  var useA = true;
  var pick = useA ? (1, 2) : (3, 4);
  var (u, v) = pick;
  print(u);
  print(v);
}"#,
        ["1", "2"]
    };

    destructure_double_values_from_function => {
        r#"(double, double) halves(double n) => (n / 2, n / 2);
void main() {
  var (a, b) = halves(10);
  print(a + b);
}"#,
        ["10.0"]
    };

    destructure_record_in_list_iteration => {
        r#"void main() {
  var pairs = [(1, 'a'), (2, 'b')];
  var chars = '';
  for (var (_, ch) in pairs) {
    chars += ch;
  }
  print(chars);
}"#,
        ["ab"]
    };

    destructure_wildcard_skips_unneeded_field => {
        r#"void main() {
  var (_, y) = (100, 200);
  print(y);
}"#,
        ["200"]
    };

    destructure_named_wildcard_keeps_other => {
        r#"void main() {
  var (name: _, id: i) = (name: 'Zed', id: 99);
  print(i);
}"#,
        ["99"]
    };

    rotate_three_strings_via_record => {
        r#"void main() {
  var a = 'x';
  var b = 'y';
  var c = 'z';
  (a, b, c) = (b, c, a);
  print(a + b + c);
}"#,
        ["yzx"]
    };

    destructure_from_method_returning_record => {
        r#"class Point {
  (int, int) coords() => (3, 4);
}
void main() {
  var (x, y) = Point().coords();
  print(x + y);
}"#,
        ["7"]
    };

    destructure_map_entry_as_record => {
        r#"void main() {
  var entry = MapEntry('k', 42);
  var (key: k, value: v) = (key: entry.key, value: entry.value);
  print(k);
  print(v);
}"#,
        ["k", "42"]
    };

    destructure_after_record_field_read => {
        r#"void main() {
  var outer = ((1, 2), label: 'pair');
  var inner = outer.$1;
  var (a, b) = inner;
  print(a);
  print(b);
}"#,
        ["1", "2"]
    };

    destructure_two_named_from_api_shape => {
        r#"void main() {
  var response = (status: 201, body: 'Created');
  var (status: s, body: b) = response;
  print(s);
  print(b);
}"#,
        ["201", "Created"]
    };

    destructure_positional_in_nested_loop => {
        r#"void main() {
  var grid = [[(1, 1), (2, 2)], [(3, 3), (4, 4)]];
  var sum = 0;
  for (var row in grid) {
    for (var (x, y) in row) {
      sum += x + y;
    }
  }
  print(sum);
}"#,
        ["20"]
    };

    destructure_record_with_nullable_named_field => {
        r#"void main() {
  var (name: n, note: t) = (name: 'Ann', note: null);
  print(n);
  print(t);
}"#,
        ["Ann", "null"]
    };

    destructure_five_tuple_from_range_like => {
        r#"void main() {
  var (a, b, c, d, e) = (1, 2, 3, 4, 5);
  print(a + b + c + d + e);
}"#,
        ["15"]
    };

    destructure_swap_preserves_sum => {
        r#"void main() {
  var x = 4;
  var y = 6;
  var before = x + y;
  (x, y) = (y, x);
  print(x + y == before);
}"#,
        ["true"]
    };

    destructure_from_closure_return => {
        r#"void main() {
  ({int x, int y}) mk() => (x: 8, y: 9);
  var (x: a, y: b) = mk();
  print(a * b);
}"#,
        ["72"]
    };

    destructure_record_with_negative_ints => {
        r#"void main() {
  var (a, b) = (-3, 5);
  print(b - a);
}"#,
        ["8"]
    };

    destructure_mixed_after_map_lookup_fallback => {
        r#"void main() {
  var m = {'p': (1, 2)};
  var (a, b) = m['p'] ?? (0, 0);
  print(a);
  print(b);
}"#,
        ["1", "2"]
    };

    destructure_three_named_from_config_record => {
        r#"void main() {
  var (host: h, port: p, tls: t) = (host: 'localhost', port: 8080, tls: false);
  print(h);
  print(p);
  print(t);
}"#,
        ["localhost", "8080", "false"]
    };

    destructure_nested_pair_from_return => {
        r#"((int, int), String) bundle() => ((9, 1), 'tag');
void main() {
  var ((a, b), tag: t) = bundle();
  print(a);
  print(b);
  print(t);
}"#,
        ["9", "1", "tag"]
    };

    destructure_final_binding_from_record => {
        r#"void main() {
  final (a, b) = (3, 4);
  print(a + b);
}"#,
        ["7"]
    };

    destructure_rebind_after_tuple_swap => {
        r#"void main() {
  var (m, n) = (1, 2);
  (m, n) = (n, m);
  var (p, q) = (m, n);
  print(p);
  print(q);
}"#,
        ["2", "1"]
    };
}
