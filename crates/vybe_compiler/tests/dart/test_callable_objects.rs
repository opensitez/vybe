//! Callable objects: call() method invocation, named and positional arguments,
//! and tear-off patterns for function-like instances.

dart_cases! {
    call_method_invoked_with_parentheses => {
        r#"class Adder {
  int call(int a, int b) {
    return a + b;
  }
}
void main() {
  var add = Adder();
  print(add(3, 4));
}"#,
        ["7"]
    };

    call_with_zero_arguments => {
        r#"class Greeter {
  String call() {
    return 'hello';
  }
}
void main() {
  print(Greeter()());
}"#,
        ["hello"]
    };

    call_with_single_positional_argument => {
        r#"class Doubler {
  int call(int n) {
    return n * 2;
  }
}
void main() {
  print(Doubler()(10));
}"#,
        ["20"]
    };

    call_with_three_positional_arguments => {
        r#"class Sum3 {
  int call(int a, int b, int c) {
    return a + b + c;
  }
}
void main() {
  print(Sum3()(1, 2, 3));
}"#,
        ["6"]
    };

    call_with_named_arguments => {
        r#"class Config {
  String call({required String mode, required int level}) {
    return '$mode:$level';
  }
}
void main() {
  print(Config()(mode: 'fast', level: 3));
}"#,
        ["fast:3"]
    };

    call_with_mixed_positional_and_named => {
        r#"class Mixer {
  int call(int base, {int bonus = 0}) {
    return base + bonus;
  }
}
void main() {
  print(Mixer()(5, bonus: 2));
}"#,
        ["7"]
    };

    call_default_named_parameter => {
        r#"class Mixer {
  int call(int base, {int bonus = 0}) {
    return base + bonus;
  }
}
void main() {
  print(Mixer()(5));
}"#,
        ["5"]
    };

    call_tear_off_via_call_property => {
        r#"class Doubler {
  int call(int n) {
    return n * 2;
  }
}
void main() {
  var d = Doubler();
  var fn = d.call;
  print(fn(6));
}"#,
        ["12"]
    };

    call_tear_off_passed_as_argument => {
        r#"int apply(int Function(int) f, int x) {
  return f(x);
}
class Tripler {
  int call(int n) {
    return n * 3;
  }
}
void main() {
  print(apply(Tripler().call, 4));
}"#,
        ["12"]
    };

    call_returns_string => {
        r#"class Formatter {
  String call(String s) {
    return s.toUpperCase();
  }
}
void main() {
  print(Formatter()('vybe'));
}"#,
        ["VYBE"]
    };

    call_returns_bool => {
        r#"class Checker {
  bool call(int n) {
    return n > 0;
  }
}
void main() {
  print(Checker()(5));
}"#,
        ["true"]
    };

    call_mutates_instance_field => {
        r#"class Counter {
  int n = 0;
  int call() {
    n++;
    return n;
  }
}
void main() {
  var c = Counter();
  c();
  print(c());
}"#,
        ["2"]
    };

    call_reads_instance_field => {
        r#"class Scaler {
  int factor = 5;
  int call(int n) {
    return n * factor;
  }
}
void main() {
  print(Scaler()(3));
}"#,
        ["15"]
    };

    call_on_subclass => {
        r#"class Base {
  int call(int n) {
    return n;
  }
}
class DoubleCall extends Base {
  @override
  int call(int n) {
    return n * 2;
  }
}
void main() {
  print(DoubleCall()(7));
}"#,
        ["14"]
    };

    call_with_optional_positional => {
        r#"class Join {
  String call([String a = 'x', String b = 'y']) {
    return a + b;
  }
}
void main() {
  print(Join()());
}"#,
        ["xy"]
    };

    call_with_one_optional_positional => {
        r#"class Join {
  String call([String a = 'x', String b = 'y']) {
    return a + b;
  }
}
void main() {
  print(Join()('a'));
}"#,
        ["ay"]
    };

    call_with_two_optional_positionals => {
        r#"class Join {
  String call([String a = 'x', String b = 'y']) {
    return a + b;
  }
}
void main() {
  print(Join()('a', 'b'));
}"#,
        ["ab"]
    };

    call_chained_on_new_expression => {
        r#"class Inc {
  int call(int n) {
    return n + 1;
  }
}
void main() {
  print(Inc()(Inc()(3)));
}"#,
        ["5"]
    };

    call_stored_in_variable_and_reused => {
        r#"class Mul {
  int call(int a, int b) {
    return a * b;
  }
}
void main() {
  var m = Mul();
  print(m(2, 3) + m(4, 5));
}"#,
        ["26"]
    };

    call_with_named_only_parameters => {
        r#"class Pair {
  String call({String left = 'L', String right = 'R'}) {
    return left + right;
  }
}
void main() {
  print(Pair()(left: 'a', right: 'b'));
}"#,
        ["ab"]
    };

    call_tear_off_invoked_twice => {
        r#"class AddOne {
  int call(int n) {
    return n + 1;
  }
}
void main() {
  var fn = AddOne().call;
  print(fn(1) + fn(2));
}"#,
        ["5"]
    };

    call_passed_to_higher_order_function => {
        r#"int twice(int Function(int) f, int x) {
  return f(f(x));
}
class Square {
  int call(int n) {
    return n * n;
  }
}
void main() {
  print(twice(Square().call, 3));
}"#,
        ["81"]
    };

    call_with_string_concatenation => {
        r#"class Appender {
  String call(String a, String b) {
    return a + '-' + b;
  }
}
void main() {
  print(Appender()('dart', 'vybe'));
}"#,
        ["dart-vybe"]
    };

    call_returns_list => {
        r#"class Range {
  List<int> call(int start, int end) {
    var out = <int>[];
    for (var i = start; i <= end; i++) {
      out.add(i);
    }
    return out;
  }
}
void main() {
  print(Range()(1, 3).join(','));
}"#,
        ["1,2,3"]
    };

    call_with_generic_like_int_behavior => {
        r#"class Max2 {
  int call(int a, int b) {
    return a > b ? a : b;
  }
}
void main() {
  print(Max2()(3, 9));
}"#,
        ["9"]
    };

    call_explicit_this_access => {
        r#"class Scale {
  int factor = 2;
  int call(int n) {
    return n * this.factor;
  }
}
void main() {
  print(Scale()(4));
}"#,
        ["8"]
    };

    call_after_field_assignment => {
        r#"class Scale {
  int factor = 1;
  int call(int n) {
    return n * factor;
  }
}
void main() {
  var s = Scale();
  s.factor = 3;
  print(s(4));
}"#,
        ["12"]
    };

    call_with_negative_numbers => {
        r#"class Diff {
  int call(int a, int b) {
    return a - b;
  }
}
void main() {
  print(Diff()(2, 5));
}"#,
        ["-3"]
    };

    call_tear_off_same_instance => {
        r#"class Id {
  int call(int n) {
    return n;
  }
}
void main() {
  var i = Id();
  var f1 = i.call;
  var f2 = i.call;
  print(f1(7) == f2(7));
}"#,
        ["true"]
    };

    call_with_multiple_named_defaults => {
        r#"class Opt {
  int call({int a = 1, int b = 2}) {
    return a + b;
  }
}
void main() {
  print(Opt()());
}"#,
        ["3"]
    };

    call_override_one_named_default => {
        r#"class Opt {
  int call({int a = 1, int b = 2}) {
    return a + b;
  }
}
void main() {
  print(Opt()(b: 10));
}"#,
        ["11"]
    };

    call_result_in_conditional_expression => {
        r#"class Sign {
  int call(int n) {
    return n >= 0 ? 1 : -1;
  }
}
void main() {
  print(Sign()(0));
}"#,
        ["1"]
    };

    call_result_used_in_expression => {
        r#"class Neg {
  int call(int n) {
    return -n;
  }
}
void main() {
  print(Neg()(5) + 10);
}"#,
        ["5"]
    };

    call_with_bool_named_args => {
        r#"class Flags {
  String call({bool on = false, bool off = true}) {
    return on.toString() + off.toString();
  }
}
void main() {
  print(Flags()(on: true, off: false));
}"#,
        ["truefalse"]
    };

    call_instance_equality_not_same_as_tear_off => {
        r#"class Fn {
  int call(int n) {
    return n;
  }
}
void main() {
  var f = Fn();
  print(f == f.call);
}"#,
        ["false"]
    };

    call_recursive_through_field => {
        r#"class Rec {
  int depth;
  Rec(this.depth);
  int call(int n) {
    if (depth <= 0) {
      return n;
    }
    return n + Rec(depth - 1)(n - 1);
  }
}
void main() {
  print(Rec(2)(3));
}"#,
        ["5"]
    };

    call_with_string_interpolation_result => {
        r#"class Tag {
  String call(String name, int id) {
    return '$name#$id';
  }
}
void main() {
  print(Tag()('item', 42));
}"#,
        ["item#42"]
    };

    call_arrow_body_style => {
        r#"class Square {
  int call(int n) => n * n;
}
void main() {
  print(Square()(6));
}"#,
        ["36"]
    };

    call_tear_off_in_list => {
        r#"class Inc {
  int call(int n) {
    return n + 1;
  }
}
void main() {
  var fns = [Inc().call, Inc().call];
  print(fns[0](2) + fns[1](3));
}"#,
        ["7"]
    };

    call_with_division => {
        r#"class Half {
  double call(int n) {
    return n / 2;
  }
}
void main() {
  print(Half()(9));
}"#,
        ["4.5"]
    };

    call_multiple_instances_independent => {
        r#"class Box {
  int scale;
  Box(this.scale);
  int call(int n) {
    return n * scale;
  }
}
void main() {
  print(Box(2)(3) + Box(5)(3));
}"#,
        ["21"]
    };
}
