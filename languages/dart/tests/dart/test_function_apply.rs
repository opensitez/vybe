//! Function.apply: positional arguments, named #arguments, and applying
//! closures, top-level functions, and instance methods dynamically.

dart_cases! {
    apply_two_positional_int_args => {
        r#"int add(int a, int b) {
  return a + b;
}
void main() {
  print(Function.apply(add, [3, 4]));
}"#,
        ["7"]
    };

    apply_three_positional_int_args => {
        r#"int sum3(int a, int b, int c) {
  return a + b + c;
}
void main() {
  print(Function.apply(sum3, [1, 2, 3]));
}"#,
        ["6"]
    };

    apply_single_positional_int_arg => {
        r#"int doubleIt(int n) {
  return n * 2;
}
void main() {
  print(Function.apply(doubleIt, [21]));
}"#,
        ["42"]
    };

    apply_zero_positional_args => {
        r#"int answer() {
  return 42;
}
void main() {
  print(Function.apply(answer, []));
}"#,
        ["42"]
    };

    apply_positional_string_concat => {
        r#"String join2(String a, String b) {
  return a + b;
}
void main() {
  print(Function.apply(join2, ['hello', ' world']));
}"#,
        ["hello world"]
    };

    apply_positional_subtraction => {
        r#"int subtract(int a, int b) {
  return a - b;
}
void main() {
  print(Function.apply(subtract, [10, 3]));
}"#,
        ["7"]
    };

    apply_positional_multiplication => {
        r#"int multiply(int a, int b) {
  return a * b;
}
void main() {
  print(Function.apply(multiply, [6, 7]));
}"#,
        ["42"]
    };

    apply_positional_integer_division => {
        r#"int divide(int a, int b) {
  return a ~/ b;
}
void main() {
  print(Function.apply(divide, [17, 5]));
}"#,
        ["3"]
    };

    apply_positional_modulo => {
        r#"int mod(int a, int b) {
  return a % b;
}
void main() {
  print(Function.apply(mod, [17, 5]));
}"#,
        ["2"]
    };

    apply_positional_bool_and => {
        r#"bool bothTrue(bool a, bool b) {
  return a && b;
}
void main() {
  print(Function.apply(bothTrue, [true, false]));
}"#,
        ["false"]
    };

    apply_positional_bool_or => {
        r#"bool eitherTrue(bool a, bool b) {
  return a || b;
}
void main() {
  print(Function.apply(eitherTrue, [true, false]));
}"#,
        ["true"]
    };

    apply_positional_compare_strings => {
        r#"bool same(String a, String b) {
  return a == b;
}
void main() {
  print(Function.apply(same, ['dart', 'dart']));
}"#,
        ["true"]
    };

    apply_positional_list_length => {
        r#"int lengthOf(List<int> items) {
  return items.length;
}
void main() {
  print(Function.apply(lengthOf, [[1, 2, 3]]));
}"#,
        ["3"]
    };

    apply_positional_on_closure => {
        r#"void main() {
  var fn = (int a, int b) => a + b;
  print(Function.apply(fn, [5, 6]));
}"#,
        ["11"]
    };

    apply_positional_on_local_function => {
        r#"void main() {
  int triple(int n) {
    return n * 3;
  }
  print(Function.apply(triple, [4]));
}"#,
        ["12"]
    };

    apply_one_required_and_one_named_arg => {
        r#"String greet(String name, {String prefix = 'Hello'}) {
  return '$prefix $name';
}
void main() {
  print(Function.apply(greet, ['Ann'], {#prefix: 'Hi'}));
}"#,
        ["Hi Ann"]
    };

    apply_named_arg_overrides_default => {
        r#"int scale(int n, {int factor = 2}) {
  return n * factor;
}
void main() {
  print(Function.apply(scale, [5], {#factor: 3}));
}"#,
        ["15"]
    };

    apply_named_arg_uses_default_when_omitted => {
        r#"int scale(int n, {int factor = 2}) {
  return n * factor;
}
void main() {
  print(Function.apply(scale, [5], {}));
}"#,
        ["10"]
    };

    apply_two_named_args => {
        r#"int compute(int a, {int b = 0, int c = 0}) {
  return a + b + c;
}
void main() {
  print(Function.apply(compute, [1], {#b: 2, #c: 3}));
}"#,
        ["6"]
    };

    apply_named_bool_flag_true => {
        r#"String label(String text, {bool upper = false}) {
  return upper ? text.toUpperCase() : text;
}
void main() {
  print(Function.apply(label, ['dart'], {#upper: true}));
}"#,
        ["DART"]
    };

    apply_named_bool_flag_false => {
        r#"String label(String text, {bool upper = false}) {
  return upper ? text.toUpperCase() : text;
}
void main() {
  print(Function.apply(label, ['dart'], {#upper: false}));
}"#,
        ["dart"]
    };

    apply_named_string_parameter => {
        r#"String repeat(String text, {String sep = ''}) {
  return text + sep + text;
}
void main() {
  print(Function.apply(repeat, ['ab'], {#sep: '-'}));
}"#,
        ["ab-ab"]
    };

    apply_positional_and_named_mixed => {
        r#"int offset(int base, {int delta = 1}) {
  return base + delta;
}
void main() {
  print(Function.apply(offset, [10], {#delta: 5}));
}"#,
        ["15"]
    };

    apply_named_only_optional_params => {
        r#"String build({String a = 'x', String b = 'y'}) {
  return a + b;
}
void main() {
  print(Function.apply(build, [], {#a: '1', #b: '2'}));
}"#,
        ["12"]
    };

    apply_positional_with_empty_named_map => {
        r#"int inc(int n, {int step = 1}) {
  return n + step;
}
void main() {
  print(Function.apply(inc, [4], {}));
}"#,
        ["5"]
    };

    apply_closure_with_named_param => {
        r#"void main() {
  var fn = (int a, {int b = 0}) => a + b;
  print(Function.apply(fn, [7], {#b: 3}));
}"#,
        ["10"]
    };

    apply_instance_method_with_positional => {
        r#"class Calc {
  int add(int a, int b) {
    return a + b;
  }
}
void main() {
  var c = Calc();
  print(Function.apply(c.add, [2, 3]));
}"#,
        ["5"]
    };

    apply_instance_method_with_named => {
        r#"class Greeter {
  String greet(String name, {String title = 'Mr'}) {
    return '$title $name';
  }
}
void main() {
  var g = Greeter();
  print(Function.apply(g.greet, ['Lee'], {#title: 'Dr'}));
}"#,
        ["Dr Lee"]
    };

    apply_static_method_positional => {
        r#"class MathUtil {
  static int max(int a, int b) {
    return a > b ? a : b;
  }
}
void main() {
  print(Function.apply(MathUtil.max, [3, 9]));
}"#,
        ["9"]
    };

    apply_function_returning_string => {
        r#"String pick(bool flag) {
  return flag ? 'yes' : 'no';
}
void main() {
  print(Function.apply(pick, [true]));
  print(Function.apply(pick, [false]));
}"#,
        ["yes", "no"]
    };

    apply_function_returning_list => {
        r#"List<int> range(int start, int end) {
  return [start, end];
}
void main() {
  var list = Function.apply(range, [1, 3]) as List;
  print(list.length);
  print(list[0]);
  print(list[1]);
}"#,
        ["2", "1", "3"]
    };

    apply_nested_apply_calls => {
        r#"int add(int a, int b) {
  return a + b;
}
void main() {
  var inner = Function.apply(add, [2, 3]);
  print(Function.apply(add, [inner, 10]));
}"#,
        ["15"]
    };

    apply_with_spread_positional_list => {
        r#"int sum3(int a, int b, int c) {
  return a + b + c;
}
void main() {
  var args = [1, 2, 3];
  print(Function.apply(sum3, args));
}"#,
        ["6"]
    };

    apply_with_variable_holding_function => {
        r#"int mul(int a, int b) {
  return a * b;
}
void main() {
  var fn = mul;
  print(Function.apply(fn, [4, 5]));
}"#,
        ["20"]
    };

    apply_positional_double_args => {
        r#"double avg(double a, double b) {
  return (a + b) / 2;
}
void main() {
  print(Function.apply(avg, [2.0, 4.0]));
}"#,
        ["3.0"]
    };

    apply_named_negative_int => {
        r#"int adjust(int n, {int delta = 0}) {
  return n + delta;
}
void main() {
  print(Function.apply(adjust, [10], {#delta: -3}));
}"#,
        ["7"]
    };

    apply_positional_string_interpolation_helper => {
        r#"String fmt(String noun, int count) {
  return '$count $noun';
}
void main() {
  print(Function.apply(fmt, ['apples', 3]));
}"#,
        ["3 apples"]
    };

    apply_named_with_multiple_positional => {
        r#"String join3(String a, String b, String c, {String sep = ''}) {
  return a + sep + b + sep + c;
}
void main() {
  print(Function.apply(join3, ['x', 'y', 'z'], {#sep: ','}));
}"#,
        ["x,y,z"]
    };

    apply_void_function_runs_print => {
        r#"void shout(String msg) {
  print(msg);
}
void main() {
  Function.apply(shout, ['hi']);
}"#,
        ["hi"]
    };

    apply_positional_on_generic_identity => {
        r#"T id<T>(T v) {
  return v;
}
void main() {
  print(Function.apply(id, [99]));
}"#,
        ["99"]
    };
}
