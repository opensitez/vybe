//! Dart 3 records: positional/named literals, field access, destructuring,
//! equality, and returning records from functions.

dart_cases! {
    positional_record_two_fields_literal => {
        r#"void main() {
  var point = (3, 4);
  print(point.$1);
  print(point.$2);
}"#,
        ["3", "4"]
    };

    positional_record_three_fields => {
        r#"void main() {
  var rgb = (255, 128, 0);
  print(rgb.$1 + rgb.$2 + rgb.$3);
}"#,
        ["383"]
    };

    named_record_single_field => {
        r#"void main() {
  var item = (count: 7);
  print(item.count);
}"#,
        ["7"]
    };

    named_record_two_fields => {
        r#"void main() {
  var user = (name: 'Ada', id: 42);
  print(user.name);
  print(user.id);
}"#,
        ["Ada", "42"]
    };

    mixed_positional_and_named_record => {
        r#"void main() {
  var entry = (1, label: 'one');
  print(entry.$1);
  print(entry.label);
}"#,
        ["1", "one"]
    };

    positional_record_dollar_one_access => {
        r#"void main() {
  var pair = ('hello', 99);
  print(pair.$1);
}"#,
        ["hello"]
    };

    positional_record_dollar_two_access => {
        r#"void main() {
  var pair = ('hello', 99);
  print(pair.$2);
}"#,
        ["99"]
    };

    named_record_field_access => {
        r#"void main() {
  var cfg = (host: 'localhost', port: 8080);
  print(cfg.host);
  print(cfg.port);
}"#,
        ["localhost", "8080"]
    };

    destructuring_positional_two_variables => {
        r#"void main() {
  var (x, y) = (10, 20);
  print(x);
  print(y);
}"#,
        ["10", "20"]
    };

    destructuring_positional_three_variables => {
        r#"void main() {
  var (a, b, c) = (1, 2, 3);
  print(a + b + c);
}"#,
        ["6"]
    };

    destructuring_named_fields => {
        r#"void main() {
  var (name: n, age: a) = (name: 'Bob', age: 30);
  print(n);
  print(a);
}"#,
        ["Bob", "30"]
    };

    destructuring_mixed_positional_and_named => {
        r#"void main() {
  var (code, message: msg) = (404, message: 'Not Found');
  print(code);
  print(msg);
}"#,
        ["404", "Not Found"]
    };

    positional_record_equality_same_values => {
        r#"void main() {
  var a = (1, 2);
  var b = (1, 2);
  print(a == b);
}"#,
        ["true"]
    };

    positional_record_inequality_different_values => {
        r#"void main() {
  var a = (1, 2);
  var b = (1, 3);
  print(a == b);
}"#,
        ["false"]
    };

    named_record_equality_same_fields => {
        r#"void main() {
  var a = (x: 5, y: 6);
  var b = (x: 5, y: 6);
  print(a == b);
}"#,
        ["true"]
    };

    named_record_inequality_different_field => {
        r#"void main() {
  var a = (x: 5, y: 6);
  var b = (x: 5, y: 7);
  print(a == b);
}"#,
        ["false"]
    };

    function_returns_positional_record => {
        r#"({int, int}) minMax(List<int> nums) {
  return (nums.first, nums.last);
}
void main() {
  var result = minMax([3, 9, 1, 7]);
  print(result.$1);
  print(result.$2);
}"#,
        ["3", "7"]
    };

    function_returns_named_record => {
        r#"({String name, int score}) topPlayer() {
  return (name: 'Zara', score: 100);
}
void main() {
  var p = topPlayer();
  print(p.name);
  print(p.score);
}"#,
        ["Zara", "100"]
    };

    function_returns_mixed_record => {
        r#"(int, {String unit}) measure() {
  return (42, unit: 'px');
}
void main() {
  var m = measure();
  print(m.$1);
  print(m.unit);
}"#,
        ["42", "px"]
    };

    nested_record_field_access => {
        r#"void main() {
  var outer = ((1, 2), label: 'pair');
  print(outer.$1.$1);
  print(outer.$1.$2);
  print(outer.label);
}"#,
        ["1", "2", "pair"]
    };

    record_stored_in_list => {
        r#"void main() {
  var points = [(0, 0), (1, 2), (3, 4)];
  print(points.length);
  print(points[1].$1);
  print(points[1].$2);
}"#,
        ["3", "1", "2"]
    };

    function_accepts_positional_record_parameter => {
        r#"int sumPair((int, int) pair) {
  return pair.$1 + pair.$2;
}
void main() {
  print(sumPair((4, 6)));
}"#,
        ["10"]
    };

    function_accepts_named_record_parameter => {
        r#"String formatUser(({String name, int id}) user) {
  return user.name + ':' + user.id.toString();
}
void main() {
  print(formatUser((name: 'Eve', id: 7)));
}"#,
        ["Eve:7"]
    };

    swap_values_via_record_destructuring => {
        r#"void main() {
  var a = 1;
  var b = 2;
  (a, b) = (b, a);
  print(a);
  print(b);
}"#,
        ["2", "1"]
    };

    single_field_positional_record => {
        r#"void main() {
  var wrap = (99,);
  print(wrap.$1);
}"#,
        ["99"]
    };

    single_field_named_record => {
        r#"void main() {
  var wrap = (value: 42);
  print(wrap.value);
}"#,
        ["42"]
    };

    destructuring_in_for_in_loop => {
        r#"void main() {
  var pairs = [(1, 'a'), (2, 'b')];
  var sum = 0;
  for (var (n, _) in pairs) {
    sum = sum + n;
  }
  print(sum);
}"#,
        ["3"]
    };

    return_record_and_destructure_immediately => {
        r#"(int, int) origin() {
  return (0, 0);
}
void main() {
  var (x, y) = origin();
  print(x);
  print(y);
}"#,
        ["0", "0"]
    };

    nested_record_equality => {
        r#"void main() {
  var a = ((1, 2), tag: 'ok');
  var b = ((1, 2), tag: 'ok');
  print(a == b);
}"#,
        ["true"]
    };

    record_with_string_fields => {
        r#"void main() {
  var pair = ('foo', 'bar');
  print(pair.$1 + pair.$2);
}"#,
        ["foobar"]
    };

    record_with_bool_and_int_fields => {
        r#"void main() {
  var status = (ok: true, code: 200);
  print(status.ok);
  print(status.code);
}"#,
        ["true", "200"]
    };

    record_passed_through_two_functions => {
        r#"(int, int) makePair() {
  return (3, 5);
}
int doubleFirst((int, int) p) {
  return p.$1 * 2;
}
void main() {
  print(doubleFirst(makePair()));
}"#,
        ["6"]
    };

    record_field_read_after_variable_assignment => {
        r#"void main() {
  var a = (10, 20);
  var b = a;
  print(b.$1);
  print(b.$2);
}"#,
        ["10", "20"]
    };

    mixed_record_equality_same_values => {
        r#"void main() {
  var a = (1, key: 'x');
  var b = (1, key: 'x');
  print(a == b);
}"#,
        ["true"]
    };

    mixed_record_inequality_on_named_part => {
        r#"void main() {
  var a = (1, key: 'x');
  var b = (1, key: 'y');
  print(a == b);
}"#,
        ["false"]
    };

    record_returned_from_arrow_function => {
        r#"({int x, int y}) point() => (x: 4, y: 8);
void main() {
  var p = point();
  print(p.x + p.y);
}"#,
        ["12"]
    };

    destructuring_named_then_reassign => {
        r#"void main() {
  var (name: n, age: a) = (name: 'Ann', age: 21);
  n = 'Ann-Lee';
  print(n);
  print(a);
}"#,
        ["Ann-Lee", "21"]
    };

    record_in_map_as_value => {
        r#"void main() {
  var map = {'a': (1, 2), 'b': (3, 4)};
  print(map['a']!.$1);
  print(map['b']!.$2);
}"#,
        ["1", "4"]
    };

    multiple_named_fields_destructured => {
        r#"void main() {
  var (r: red, g: green, b: blue) = (r: 1, g: 2, b: 3);
  print(red + green + blue);
}"#,
        ["6"]
    };

    record_type_with_double_values => {
        r#"void main() {
  (double, double) scale = (1.5, 2.5);
  print(scale.$1 + scale.$2);
}"#,
        ["4.0"]
    };
}
