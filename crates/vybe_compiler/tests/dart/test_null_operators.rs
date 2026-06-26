//! Null-aware operators: ?., ??, ??=, !, late initialization, and required
//! named parameters interacting with nullable types.

dart_cases! {
    null_aware_property_access_on_null_returns_null => {
        r#"class Box { int? value; }
void main() {
  var b = Box();
  print(b.value?.toString() ?? 'missing');
}"#,
        ["missing"]
    };

    null_aware_property_access_on_value_calls_method => {
        r#"class Box { int? value = 7; }
void main() {
  var b = Box();
  print(b.value?.isEven);
}"#,
        ["false"]
    };

    null_aware_method_chain_short_circuits_on_null => {
        r#"class Inner { String label = 'ok'; }
class Outer { Inner? inner; }
void main() {
  var o = Outer();
  print(o.inner?.label ?? 'none');
}"#,
        ["none"]
    };

    null_aware_method_chain_follows_through_when_present => {
        r#"class Inner { String label = 'ok'; }
class Outer { Inner? inner = Inner(); }
void main() {
  var o = Outer();
  print(o.inner?.label);
}"#,
        ["ok"]
    };

    null_aware_index_on_nullable_list_element => {
        r#"void main() {
  List<int?> nums = [1, null, 3];
  print(nums[1]?.toString() ?? 'null-slot');
}"#,
        ["null-slot"]
    };

    null_aware_cascade_on_nullable_receiver => {
        r#"class Counter { int n = 0; void inc() { n++; } }
void main() {
  Counter? c;
  c?.inc();
  print(c?.n ?? -1);
}"#,
        ["-1"]
    };

    null_aware_cascade_mutates_when_receiver_present => {
        r#"class Counter { int n = 0; void inc() { n++; } }
void main() {
  Counter? c = Counter();
  c?.inc();
  print(c?.n);
}"#,
        ["1"]
    };

    coalesce_returns_left_when_non_null_int => {
        r#"void main() {
  int? a = 5;
  print(a ?? 99);
}"#,
        ["5"]
    };

    coalesce_returns_right_when_left_is_null => {
        r#"void main() {
  int? a;
  print(a ?? 99);
}"#,
        ["99"]
    };

    coalesce_chained_skips_all_nulls => {
        r#"void main() {
  String? a;
  String? b;
  String? c = 'found';
  print(a ?? b ?? c ?? 'fallback');
}"#,
        ["found"]
    };

    coalesce_chained_reaches_final_fallback => {
        r#"void main() {
  String? a;
  String? b;
  print(a ?? b ?? 'fallback');
}"#,
        ["fallback"]
    };

    coalesce_with_arithmetic_expression => {
        r#"void main() {
  int? n;
  print((n ?? 3) * 4);
}"#,
        ["12"]
    };

    coalesce_preserves_empty_string_as_non_null => {
        r#"void main() {
  String? s = '';
  print(s ?? 'default');
}"#,
        [""]
    };

    coalesce_preserves_zero_as_non_null => {
        r#"void main() {
  int? n = 0;
  print(n ?? 42);
}"#,
        ["0"]
    };

    coalesce_with_bool_nullable => {
        r#"void main() {
  bool? flag;
  print(flag ?? true);
}"#,
        ["true"]
    };

    coalesce_with_double_nullable => {
        r#"void main() {
  double? d;
  print(d ?? 2.5);
}"#,
        ["2.5"]
    };

    null_assign_sets_when_variable_is_null => {
        r#"void main() {
  String? s;
  s ??= 'assigned';
  print(s);
}"#,
        ["assigned"]
    };

    null_assign_preserves_existing_non_null_value => {
        r#"void main() {
  String? s = 'keep';
  s ??= 'replace';
  print(s);
}"#,
        ["keep"]
    };

    null_assign_on_int_nullable => {
        r#"void main() {
  int? count;
  count ??= 10;
  print(count);
}"#,
        ["10"]
    };

    null_assign_chained_on_field => {
        r#"class Cache {
  String? token;
  String ensureToken() {
    token ??= 'generated';
    return token!;
  }
}
void main() {
  var c = Cache();
  print(c.ensureToken());
  print(c.ensureToken());
}"#,
        ["generated", "generated"]
    };

    null_assign_does_not_overwrite_zero => {
        r#"void main() {
  int? n = 0;
  n ??= 99;
  print(n);
}"#,
        ["0"]
    };

    non_null_assert_on_string_promotes_value => {
        r#"void main() {
  String? s = 'dart';
  print(s!.length);
}"#,
        ["4"]
    };

    non_null_assert_on_int_in_addition => {
        r#"void main() {
  int? n = 8;
  print(n! + 2);
}"#,
        ["10"]
    };

    non_null_assert_on_bool => {
        r#"void main() {
  bool? flag = false;
  print(flag!);
}"#,
        ["false"]
    };

    non_null_assert_on_list => {
        r#"void main() {
  List<int>? items = [1, 2];
  print(items!.length);
}"#,
        ["2"]
    };

    non_null_assert_in_return_expression => {
        r#"String? pick(String? a, String? b) => a ?? b!;
void main() {
  print(pick(null, 'second'));
}"#,
        ["second"]
    };

    late_variable_assigned_before_read => {
        r#"late int total;
void main() {
  total = 100;
  print(total);
}"#,
        ["100"]
    };

    late_string_assigned_in_main => {
        r#"late String name;
void main() {
  name = 'vybe';
  print(name);
}"#,
        ["vybe"]
    };

    late_final_assigned_once => {
        r#"late final int seed;
void main() {
  seed = 7;
  print(seed);
}"#,
        ["7"]
    };

    late_field_in_class_initialized_on_access => {
        r#"class Holder {
  late String tag;
  Holder() { tag = 'ready'; }
}
void main() {
  print(Holder().tag);
}"#,
        ["ready"]
    };

    late_nullable_stores_null => {
        r#"late String? maybe;
void main() {
  maybe = null;
  print(maybe ?? 'empty');
}"#,
        ["empty"]
    };

    late_nullable_stores_value => {
        r#"late String? maybe;
void main() {
  maybe = 'set';
  print(maybe ?? 'empty');
}"#,
        ["set"]
    };

    required_named_param_accepts_explicit_null => {
        r#"void log({required String? message}) {
  print(message ?? 'null-message');
}
void main() {
  log(message: null);
}"#,
        ["null-message"]
    };

    required_named_param_accepts_non_null => {
        r#"void log({required String? message}) {
  print(message ?? 'null-message');
}
void main() {
  log(message: 'hello');
}"#,
        ["hello"]
    };

    required_named_int_param_with_nullable_type => {
        r#"void show({required int? count}) {
  print(count ?? 0);
}
void main() {
  show(count: null);
}"#,
        ["0"]
    };

    required_named_bool_param_with_null_coalesce => {
        r#"void flag({required bool? enabled}) {
  print(enabled ?? false);
}
void main() {
  flag(enabled: true);
}"#,
        ["true"]
    };

    nullable_required_named_mixed_with_optional_default => {
        r#"void connect({required String? host, int port = 8080}) {
  print('${host ?? 'localhost'}:$port');
}
void main() {
  connect(host: null);
}"#,
        ["localhost:8080"]
    };

    null_aware_call_on_nullable_function_variable => {
        r#"void main() {
  int Function(int)? doubler;
  print(doubler?.call(5) ?? -1);
}"#,
        ["-1"]
    };

    null_aware_call_on_present_function_variable => {
        r#"void main() {
  int Function(int)? doubler = (x) => x * 2;
  print(doubler?.call(5));
}"#,
        ["10"]
    };

    null_aware_map_lookup_on_nullable_map => {
        r#"void main() {
  Map<String, int>? scores;
  print(scores?['alice'] ?? 0);
}"#,
        ["0"]
    };

    null_aware_map_lookup_on_present_map => {
        r#"void main() {
  Map<String, int>? scores = {'alice': 10};
  print(scores?['alice']);
}"#,
        ["10"]
    };

    coalesce_nested_with_null_aware_access => {
        r#"class Node { String? name; Node? next; }
void main() {
  var n = Node();
  print(n.next?.name ?? n.name ?? 'anon');
}"#,
        ["anon"]
    };
}
