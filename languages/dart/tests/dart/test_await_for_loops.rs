//! Dart await for loops over async* streams — break, continue, nesting, edge cases.

dart_cases! {
    await_for_reads_async_star_stream_in_order => {
        r#"Stream<int> nums() async* { yield 1; yield 2; yield 3; }
Future<void> main() async {
  var out = <int>[];
  await for (var n in nums()) { out.add(n); }
  print(out.join(','));
}"#,
        ["1,2,3"]
    };

    await_for_on_empty_async_star_runs_zero_times => {
        r#"Stream<int> empty() async* {}
Future<void> main() async {
  var count = 0;
  await for (var _ in empty()) { count++; }
  print(count);
}"#,
        ["0"]
    };

    await_for_on_single_element_async_star => {
        r#"Stream<int> solo() async* { yield 42; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in solo()) { out.add(v); }
  print(out.join('|'));
}"#,
        ["42"]
    };

    await_for_prints_each_event_directly => {
        r#"Stream<int> gen() async* { yield 5; yield 6; }
Future<void> main() async {
  await for (var v in gen()) { print(v); }
}"#,
        ["5", "6"]
    };

    await_for_break_exits_before_remaining_events => {
        r#"Stream<int> gen() async* { for (var i = 1; i <= 6; i++) yield i; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in gen()) {
    if (v == 4) break;
    out.add(v);
  }
  print(out.join(','));
}"#,
        ["1,2,3"]
    };

    await_for_continue_skips_selected_events => {
        r#"Stream<int> gen() async* { for (var i = 1; i <= 5; i++) yield i; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in gen()) {
    if (v % 2 == 0) continue;
    out.add(v);
  }
  print(out.join(','));
}"#,
        ["1,3,5"]
    };

    await_for_break_on_first_iteration => {
        r#"Stream<int> gen() async* { yield 9; yield 10; yield 11; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in gen()) {
    out.add(v);
    break;
  }
  print(out.join(','));
}"#,
        ["9"]
    };

    await_for_continue_on_all_iterations_yields_nothing => {
        r#"Stream<int> gen() async* { yield 1; yield 2; yield 3; }
Future<void> main() async {
  var count = 0;
  await for (var _ in gen()) { continue; count++; }
  print(count);
}"#,
        ["0"]
    };

    await_for_nested_async_star_inner_then_outer => {
        r#"Stream<int> inner() async* { yield 2; yield 3; }
Stream<int> outer() async* { yield 1; yield* inner(); yield 4; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in outer()) { out.add(v); }
  print(out.join(','));
}"#,
        ["1,2,3,4"]
    };

    await_for_nested_loops_over_two_generators => {
        r#"Stream<int> rows() async* { yield 1; yield 2; }
Stream<int> cols() async* { yield 10; yield 20; }
Future<void> main() async {
  var out = <String>[];
  await for (var r in rows()) {
    await for (var c in cols()) {
      out.add('$r$c');
    }
  }
  print(out.join(','));
}"#,
        ["110,120,210,220"]
    };

    await_for_inner_break_does_not_stop_outer => {
        r#"Stream<int> outer() async* { yield 1; yield 2; }
Stream<int> inner() async* { yield 1; yield 2; yield 3; }
Future<void> main() async {
  var out = <int>[];
  await for (var o in outer()) {
    await for (var i in inner()) {
      if (i == 2) break;
      out.add(o * 10 + i);
    }
  }
  print(out.join(','));
}"#,
        ["11,21"]
    };

    await_for_outer_break_stops_both_levels => {
        r#"Stream<int> outer() async* { yield 1; yield 2; yield 3; }
Stream<int> inner() async* { yield 10; yield 20; }
Future<void> main() async {
  var out = <int>[];
  await for (var o in outer()) {
    await for (var i in inner()) {
      out.add(o + i);
    }
    if (o == 1) break;
  }
  print(out.join(','));
}"#,
        ["11,21"]
    };

    await_for_accumulates_sum_from_async_star => {
        r#"Stream<int> gen() async* { yield 1; yield 2; yield 3; yield 4; }
Future<void> main() async {
  var sum = 0;
  await for (var v in gen()) { sum += v; }
  print(sum);
}"#,
        ["10"]
    };

    await_for_with_string_async_star => {
        r#"Stream<String> gen() async* { yield 'a'; yield 'b'; yield 'c'; }
Future<void> main() async {
  print(await gen().join(''));
}"#,
        ["abc"]
    };

    await_for_over_yield_star_delegated_stream => {
        r#"Stream<int> inner() async* { yield 7; yield 8; }
Stream<int> outer() async* { yield* inner(); }
Future<void> main() async {
  var out = <int>[];
  await for (var v in outer()) { out.add(v); }
  print(out.join(','));
}"#,
        ["7,8"]
    };

    await_for_after_await_in_generator => {
        r#"Future<int> ready() async => 0;
Stream<int> gen() async* {
  var start = await ready();
  yield start + 1;
  yield start + 2;
}
Future<void> main() async {
  var out = <int>[];
  await for (var v in gen()) { out.add(v); }
  print(out.join(','));
}"#,
        ["1,2"]
    };

    await_for_empty_stream_leaves_accumulator_unchanged => {
        r#"Stream<int> empty() async* {}
Future<void> main() async {
  var sum = 5;
  await for (var v in empty()) { sum += v; }
  print(sum);
}"#,
        ["5"]
    };

    await_for_single_element_updates_counter_once => {
        r#"Stream<int> one() async* { yield 100; }
Future<void> main() async {
  var hits = 0;
  await for (var _ in one()) { hits++; }
  print(hits);
}"#,
        ["1"]
    };

    await_for_break_with_label_like_early_exit_pattern => {
        r#"Stream<int> gen() async* { for (var i = 0; i < 10; i++) yield i; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in gen()) {
    if (v >= 3) break;
    out.add(v);
  }
  print(out.join(','));
}"#,
        ["0,1,2"]
    };

    await_for_continue_filters_multiples_of_three => {
        r#"Stream<int> gen() async* { for (var i = 1; i <= 9; i++) yield i; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in gen()) {
    if (v % 3 == 0) continue;
    out.add(v);
  }
  print(out.join(','));
}"#,
        ["1,2,4,5,7,8"]
    };

    await_for_on_async_star_with_await_between_yields => {
        r#"Future<int> bump(int n) async => n + 1;
Stream<int> gen() async* {
  yield await bump(0);
  yield await bump(1);
}
Future<void> main() async {
  var out = <int>[];
  await for (var v in gen()) { out.add(v); }
  print(out.join(','));
}"#,
        ["1,2"]
    };

    await_for_catches_error_from_async_star => {
        r#"Stream<int> bad() async* {
  yield 1;
  throw Exception('fail');
}
Future<void> main() async {
  var out = <String>[];
  try {
    await for (var v in bad()) { out.add('$v'); }
  } catch (_) {
    out.add('caught');
  }
  print(out.join(','));
}"#,
        ["1,caught"]
    };

    await_for_error_before_first_event => {
        r#"Stream<int> fail() async* { throw Exception('x'); yield 1; }
Future<void> main() async {
  var ok = false;
  try {
    await for (var _ in fail()) { ok = true; }
  } catch (_) {}
  print(ok);
}"#,
        ["false"]
    };

    await_for_over_recursive_async_star => {
        r#"Stream<int> down(int n) async* {
  if (n <= 0) return;
  yield n;
  yield* down(n - 1);
}
Future<void> main() async {
  var out = <int>[];
  await for (var v in down(3)) { out.add(v); }
  print(out.join(','));
}"#,
        ["3,2,1"]
    };

    await_for_fibonacci_take_five => {
        r#"Stream<int> fib() async* {
  var a = 0, b = 1;
  while (true) { yield a; var c = a + b; a = b; b = c; }
}
Future<void> main() async {
  var out = <int>[];
  await for (var v in fib().take(5)) { out.add(v); }
  print(out.join(','));
}"#,
        ["0,1,1,2,3"]
    };

    await_for_with_if_inside_body => {
        r#"Stream<int> gen() async* { yield 1; yield 2; yield 3; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in gen()) {
    if (v > 1) { out.add(v * 10); }
  }
  print(out.join(','));
}"#,
        ["20,30"]
    };

    await_for_nested_empty_inner_stream => {
        r#"Stream<int> outer() async* { yield 1; yield 2; }
Stream<int> inner() async* {}
Future<void> main() async {
  var out = <int>[];
  await for (var o in outer()) {
    await for (var i in inner()) { out.add(o + i); }
    out.add(o);
  }
  print(out.join(','));
}"#,
        ["1,2"]
    };

    await_for_nested_single_inner_event => {
        r#"Stream<int> outer() async* { yield 1; }
Stream<int> inner() async* { yield 9; }
Future<void> main() async {
  var out = <int>[];
  await for (var o in outer()) {
    await for (var i in inner()) { out.add(o + i); }
  }
  print(out.join(','));
}"#,
        ["10"]
    };

    await_for_break_after_single_print => {
        r#"Stream<int> gen() async* { yield 7; yield 8; yield 9; }
Future<void> main() async {
  await for (var v in gen()) {
    print(v);
    break;
  }
}"#,
        ["7"]
    };

    await_for_continue_then_break_combination => {
        r#"Stream<int> gen() async* { for (var i = 0; i < 6; i++) yield i; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in gen()) {
    if (v == 0) continue;
    if (v == 5) break;
    out.add(v);
  }
  print(out.join(','));
}"#,
        ["1,2,3,4"]
    };

    await_for_on_async_star_map_transform => {
        r#"Stream<int> gen() async* { yield 1; yield 2; yield 3; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in gen().map((x) => x + 1)) { out.add(v); }
  print(out.join(','));
}"#,
        ["2,3,4"]
    };

    await_for_on_async_star_where_filter => {
        r#"Stream<int> gen() async* { for (var i = 1; i <= 6; i++) yield i; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in gen().where((x) => x % 2 == 0)) { out.add(v); }
  print(out.join(','));
}"#,
        ["2,4,6"]
    };

    await_for_on_async_star_take_limited => {
        r#"Stream<int> gen() async* { for (var i = 1; i <= 10; i++) yield i; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in gen().take(3)) { out.add(v); }
  print(out.join(','));
}"#,
        ["1,2,3"]
    };

    await_for_on_async_star_skip_prefix => {
        r#"Stream<int> gen() async* { for (var i = 1; i <= 5; i++) yield i; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in gen().skip(3)) { out.add(v); }
  print(out.join(','));
}"#,
        ["4,5"]
    };

    await_for_three_nested_generators_flattened => {
        r#"Stream<int> a() async* { yield 1; }
Stream<int> b() async* { yield 2; }
Stream<int> c() async* { yield 3; }
Future<void> main() async {
  var out = <int>[];
  await for (var x in a()) {
    await for (var y in b()) {
      await for (var z in c()) {
        out.add(x + y + z);
      }
    }
  }
  print(out.join(','));
}"#,
        ["6"]
    };

    await_for_string_join_from_events => {
        r#"Stream<String> gen() async* { yield 'hello'; yield ' '; yield 'dart'; }
Future<void> main() async {
  var buf = '';
  await for (var s in gen()) { buf = buf + s; }
  print(buf);
}"#,
        ["hello dart"]
    };

    await_for_bool_stream_prints_flags => {
        r#"Stream<bool> gen() async* { yield true; yield false; yield true; }
Future<void> main() async {
  var out = <String>[];
  await for (var b in gen()) { out.add('$b'); }
  print(out.join('|'));
}"#,
        ["true|false|true"]
    };

    await_for_over_async_star_with_trailing_yield_after_delegation => {
        r#"Stream<int> inner() async* { yield 2; }
Stream<int> outer() async* { yield 1; yield* inner(); yield 3; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in outer()) { out.add(v); }
  print(out.join(','));
}"#,
        ["1,2,3"]
    };

    await_for_loop_variable_shadowing_ok => {
        r#"Stream<int> gen() async* { yield 1; yield 2; }
Future<void> main() async {
  var v = 99;
  await for (var v in gen()) { print(v); }
  print(v);
}"#,
        ["1", "2", "99"]
    };

    await_for_zero_events_with_break_immediately => {
        r#"Stream<int> empty() async* {}
Future<void> main() async {
  var ran = false;
  await for (var _ in empty()) { ran = true; break; }
  print(ran);
}"#,
        ["false"]
    };

    await_for_single_event_then_break => {
        r#"Stream<int> gen() async* { yield 55; yield 66; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in gen()) {
    out.add(v);
    if (out.length == 1) break;
  }
  print(out.join(','));
}"#,
        ["55"]
    };

    await_for_over_infinite_generator_with_take => {
        r#"Stream<int> naturals() async* {
  var n = 0;
  while (true) { yield n; n++; }
}
Future<void> main() async {
  var out = <int>[];
  await for (var v in naturals().take(4)) { out.add(v); }
  print(out.join(','));
}"#,
        ["0,1,2,3"]
    };

    await_for_with_local_counter_in_body => {
        r#"Stream<int> gen() async* { yield 3; yield 4; yield 5; }
Future<void> main() async {
  var idx = 0;
  await for (var v in gen()) {
    print('$idx:$v');
    idx++;
  }
}"#,
        ["0:3", "1:4", "2:5"]
    };

    await_for_on_parameterized_async_star => {
        r#"Stream<int> repeat(int n, int times) async* {
  for (var i = 0; i < times; i++) { yield n; }
}
Future<void> main() async {
  var out = <int>[];
  await for (var v in repeat(7, 3)) { out.add(v); }
  print(out.join(','));
}"#,
        ["7,7,7"]
    };

    await_for_sequential_two_generators => {
        r#"Stream<int> first() async* { yield 1; yield 2; }
Stream<int> second() async* { yield 3; yield 4; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in first()) { out.add(v); }
  await for (var v in second()) { out.add(v); }
  print(out.join(','));
}"#,
        ["1,2,3,4"]
    };

}
