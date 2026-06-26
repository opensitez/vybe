//! Dart async* generators — yield, yield*, await-before-yield, Stream semantics.

dart_cases! {
    async_star_yields_single_value_via_await_for => {
        r#"Stream<int> one() async* { yield 1; }
Future<void> main() async {
  await for (var v in one()) { print(v); }
}"#,
        ["1"]
    };

    async_star_yields_two_values_in_order => {
        r#"Stream<int> pair() async* { yield 10; yield 20; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in pair()) { out.add(v); }
  print(out.join(','));
}"#,
        ["10,20"]
    };

    async_star_yields_three_strings => {
        r#"Stream<String> tags() async* { yield 'x'; yield 'y'; yield 'z'; }
Future<void> main() async {
  print(await tags().join(''));
}"#,
        ["xyz"]
    };

    async_star_empty_stream_yields_nothing => {
        r#"Stream<int> empty() async* {}
Future<void> main() async {
  var n = 0;
  await for (var _ in empty()) { n++; }
  print(n);
}"#,
        ["0"]
    };

    async_star_single_element_stream => {
        r#"Stream<int> solo() async* { yield 99; }
Future<void> main() async {
  print(await solo().length);
}"#,
        ["1"]
    };

    async_star_range_loop_yields_sequence => {
        r#"Stream<int> range(int n) async* {
  for (var i = 0; i < n; i++) { yield i; }
}
Future<void> main() async {
  print(await range(5).join(','));
}"#,
        ["0,1,2,3,4"]
    };

    async_star_while_loop_yields_countdown => {
        r#"Stream<int> down(int n) async* {
  while (n > 0) { yield n; n--; }
}
Future<void> main() async {
  print(await down(3).join(','));
}"#,
        ["3,2,1"]
    };

    async_star_await_before_first_yield => {
        r#"Future<int> bump(int n) async => n + 1;
Stream<int> delayed() async* {
  var v = await bump(4);
  yield v;
  yield v + 1;
}
Future<void> main() async {
  print(await delayed().join(','));
}"#,
        ["5,6"]
    };

    async_star_await_between_yields => {
        r#"Future<int> id(int n) async => n;
Stream<int> spaced() async* {
  yield await id(1);
  yield await id(2);
  yield await id(3);
}
Future<void> main() async {
  print(await spaced().join(','));
}"#,
        ["1,2,3"]
    };

    async_star_await_future_value_then_yield => {
        r#"Future<String> label() async => 'ok';
Stream<String> tagged() async* {
  yield await label();
}
Future<void> main() async {
  print(await tagged().first);
}"#,
        ["ok"]
    };

    async_star_yield_star_delegates_inner_async_stream => {
        r#"Stream<int> inner() async* { yield 1; yield 2; }
Stream<int> outer() async* { yield* inner(); }
Future<void> main() async {
  print(await outer().join(','));
}"#,
        ["1,2"]
    };

    async_star_yield_star_then_trailing_yield => {
        r#"Stream<int> inner() async* { yield 1; }
Stream<int> outer() async* { yield* inner(); yield 2; }
Future<void> main() async {
  print(await outer().join(','));
}"#,
        ["1,2"]
    };

    async_star_yield_star_from_sync_iterable => {
        r#"Stream<int> fromList() async* { yield* [3, 4, 5]; }
Future<void> main() async {
  print(await fromList().join(','));
}"#,
        ["3,4,5"]
    };

    async_star_yield_star_from_string => {
        r#"Stream<String> chars() async* { yield* 'ab'; }
Future<void> main() async {
  print(await chars().join(''));
}"#,
        ["ab"]
    };

    async_star_nested_yield_star_chain => {
        r#"Stream<int> a() async* { yield 1; }
Stream<int> b() async* { yield* a(); yield 2; }
Stream<int> c() async* { yield* b(); yield 3; }
Future<void> main() async {
  print(await c().join(','));
}"#,
        ["1,2,3"]
    };

    async_star_multiple_await_before_multiple_yields => {
        r#"Future<int> step(int n) async => n * 2;
Stream<int> steps() async* {
  yield await step(1);
  yield await step(2);
}
Future<void> main() async {
  print(await steps().join(','));
}"#,
        ["2,4"]
    };

    async_star_if_branch_controls_yield => {
        r#"Stream<int> pick(bool useA) async* {
  if (useA) { yield 7; } else { yield 8; }
}
Future<void> main() async {
  print(await pick(true).first);
}"#,
        ["7"]
    };

    async_star_continue_skips_even_yields => {
        r#"Stream<int> odds(int n) async* {
  for (var i = 0; i < n; i++) {
    if (i % 2 == 0) continue;
    yield i;
  }
}
Future<void> main() async {
  print(await odds(6).join(','));
}"#,
        ["1,3,5"]
    };

    async_star_break_limits_loop_yields => {
        r#"Stream<int> capped() async* {
  for (var i = 0; i < 10; i++) {
    if (i == 3) break;
    yield i;
  }
}
Future<void> main() async {
  print(await capped().join(','));
}"#,
        ["0,1,2"]
    };

    async_star_recursive_async_generator => {
        r#"Stream<int> down(int n) async* {
  if (n <= 0) return;
  yield n;
  yield* down(n - 1);
}
Future<void> main() async {
  print(await down(3).join(','));
}"#,
        ["3,2,1"]
    };

    async_star_fibonacci_take_seven => {
        r#"Stream<int> fib() async* {
  var a = 0, b = 1;
  while (true) { yield a; var c = a + b; a = b; b = c; }
}
Future<void> main() async {
  print(await fib().take(7).join(','));
}"#,
        ["0,1,1,2,3,5,8"]
    };

    async_star_map_on_stream_doubles => {
        r#"Stream<int> gen() async* { yield 1; yield 2; yield 3; }
Future<void> main() async {
  print(await gen().map((x) => x * 3).join(','));
}"#,
        ["3,6,9"]
    };

    async_star_where_filters_stream => {
        r#"Stream<int> gen() async* { for (var i = 1; i <= 6; i++) yield i; }
Future<void> main() async {
  print(await gen().where((x) => x % 2 == 0).join(','));
}"#,
        ["2,4,6"]
    };

    async_star_take_limits_infinite_style => {
        r#"Stream<int> naturals() async* {
  var n = 1;
  while (true) { yield n; n++; }
}
Future<void> main() async {
  print(await naturals().take(4).join(','));
}"#,
        ["1,2,3,4"]
    };

    async_star_skip_then_take => {
        r#"Stream<int> seq() async* { for (var i = 0; i < 6; i++) yield i; }
Future<void> main() async {
  print(await seq().skip(1).take(3).join(','));
}"#,
        ["1,2,3"]
    };

    async_star_to_list_materializes_all_yields => {
        r#"Stream<int> gen() async* { yield 4; yield 5; yield 6; }
Future<void> main() async {
  var list = await gen().toList();
  print(list.join('|'));
}"#,
        ["4|5|6"]
    };

    async_star_first_and_last_selectors => {
        r#"Stream<int> gen() async* { yield 10; yield 20; yield 30; }
Future<void> main() async {
  print(await gen().first);
  print(await gen().last);
}"#,
        ["10", "30"]
    };

    async_star_fold_accumulates_with_seed => {
        r#"Stream<int> gen() async* { yield 1; yield 2; yield 3; }
Future<void> main() async {
  print(await gen().fold(10, (a, b) => a + b));
}"#,
        ["16"]
    };

    async_star_reduce_multiplies => {
        r#"Stream<int> gen() async* { yield 2; yield 3; }
Future<void> main() async {
  print(await gen().reduce((a, b) => a * b));
}"#,
        ["6"]
    };

    async_star_contains_checks_events => {
        r#"Stream<int> gen() async* { yield 5; yield 15; }
Future<void> main() async {
  print(await gen().contains(15));
}"#,
        ["true"]
    };

    async_star_length_counts_events => {
        r#"Stream<int> gen() async* { yield 1; yield 2; yield 3; yield 4; }
Future<void> main() async {
  print(await gen().length);
}"#,
        ["4"]
    };

    async_star_is_empty_false_for_nonempty => {
        r#"Stream<int> gen() async* { yield 0; }
Future<void> main() async {
  print(await gen().isEmpty);
}"#,
        ["false"]
    };

    async_star_is_empty_true_for_empty_generator => {
        r#"Stream<int> gen() async* {}
Future<void> main() async {
  print(await gen().isEmpty);
}"#,
        ["true"]
    };

    async_star_error_thrown_after_yield_propagates => {
        r#"Stream<int> bad() async* {
  yield 1;
  throw Exception('boom');
}
Future<void> main() async {
  var out = <String>[];
  try {
    await for (var v in bad()) { out.add('$v'); }
  } catch (e) {
    out.add('err');
  }
  print(out.join(','));
}"#,
        ["1,err"]
    };

    async_star_error_before_any_yield_reaches_consumer => {
        r#"Stream<int> failEarly() async* {
  throw Exception('early');
  yield 1;
}
Future<void> main() async {
  var caught = false;
  try {
    await for (var _ in failEarly()) {}
  } catch (_) {
    caught = true;
  }
  print(caught);
}"#,
        ["true"]
    };

    async_star_yield_star_from_async_inner_with_await => {
        r#"Stream<int> inner() async* {
  yield await Future.value(2);
  yield 3;
}
Stream<int> outer() async* { yield 1; yield* inner(); }
Future<void> main() async {
  print(await outer().join(','));
}"#,
        ["1,2,3"]
    };

    async_star_await_for_consumes_nested_async_star => {
        r#"Stream<int> inner() async* { yield 4; yield 5; }
Stream<int> outer() async* { yield* inner(); yield 6; }
Future<void> main() async {
  var out = <int>[];
  await for (var v in outer()) { out.add(v); }
  print(out.join(','));
}"#,
        ["4,5,6"]
    };

    async_star_local_state_between_yields => {
        r#"Stream<int> counter() async* {
  var n = 0;
  yield n;
  n++;
  yield n;
  n++;
  yield n;
}
Future<void> main() async {
  print(await counter().join(','));
}"#,
        ["0,1,2"]
    };

    async_star_parameterized_stream => {
        r#"Stream<int> repeat(int v, int times) async* {
  for (var i = 0; i < times; i++) { yield v; }
}
Future<void> main() async {
  print(await repeat(8, 3).join(','));
}"#,
        ["8,8,8"]
    };

    async_star_expand_flattens_nested_lists => {
        r#"Stream<List<int>> gen() async* {
  yield [1, 2];
  yield [3];
}
Future<void> main() async {
  print(await gen().expand((p) => p).join(','));
}"#,
        ["1,2,3"]
    };

    async_star_distinct_removes_adjacent_dupes => {
        r#"Stream<int> gen() async* { yield 1; yield 1; yield 2; yield 2; yield 2; yield 3; }
Future<void> main() async {
  print(await gen().distinct().join(','));
}"#,
        ["1,2,3"]
    };

    async_star_take_while_stops_at_predicate => {
        r#"Stream<int> gen() async* { for (var i = 1; i <= 6; i++) yield i; }
Future<void> main() async {
  print(await gen().takeWhile((x) => x < 4).join(','));
}"#,
        ["1,2,3"]
    };

    async_star_skip_while_skips_prefix => {
        r#"Stream<int> gen() async* { yield 0; yield 0; yield 1; yield 2; }
Future<void> main() async {
  print(await gen().skipWhile((x) => x == 0).join(','));
}"#,
        ["1,2"]
    };

    async_star_for_each_runs_per_event => {
        r#"Stream<int> gen() async* { yield 1; yield 2; yield 3; }
Future<void> main() async {
  var sum = 0;
  await gen().forEach((v) => sum += v);
  print(sum);
}"#,
        ["6"]
    };

    async_star_every_checks_all_events => {
        r#"Stream<int> gen() async* { yield 2; yield 4; yield 6; }
Future<void> main() async {
  print(await gen().every((x) => x % 2 == 0));
}"#,
        ["true"]
    };

    async_star_any_finds_large_value => {
        r#"Stream<int> gen() async* { yield 1; yield 2; yield 9; }
Future<void> main() async {
  print(await gen().any((x) => x > 5));
}"#,
        ["true"]
    };

    async_star_element_at_reads_position => {
        r#"Stream<String> gen() async* { yield 'p'; yield 'q'; yield 'r'; }
Future<void> main() async {
  print(await gen().elementAt(1));
}"#,
        ["q"]
    };

    async_star_await_in_loop_before_yield => {
        r#"Future<int> next(int n) async => n + 1;
Stream<int> pipeline(int count) async* {
  var v = 0;
  for (var i = 0; i < count; i++) {
    v = await next(v);
    yield v;
  }
}
Future<void> main() async {
  print(await pipeline(3).join(','));
}"#,
        ["1,2,3"]
    };

    async_star_yield_bool_values => {
        r#"Stream<bool> flags() async* { yield true; yield false; }
Future<void> main() async {
  var out = <String>[];
  await for (var f in flags()) { out.add('$f'); }
  print(out.join(','));
}"#,
        ["true,false"]
    };

    async_star_try_finally_runs_on_completion => {
        r#"Stream<int> gen() async* {
  try { yield 1; yield 2; } finally { print('done'); }
}
Future<void> main() async {
  print(await gen().join(','));
}"#,
        ["done", "1,2"]
    };

    async_star_multiple_streams_independent => {
        r#"Stream<int> gen() async* { yield 1; yield 2; }
Future<void> main() async {
  print(await gen().first);
  print(await gen().first);
}"#,
        ["1", "1"]
    };

    async_star_async_map_with_await => {
        r#"Stream<int> gen() async* { yield 1; yield 2; }
Future<void> main() async {
  print(await gen().asyncMap((x) async => x + 5).join(','));
}"#,
        ["6,7"]
    };

    async_star_chain_map_then_where => {
        r#"Stream<int> gen() async* { for (var i = 1; i <= 5; i++) yield i; }
Future<void> main() async {
  var s = gen().map((x) => x * 2).where((x) => x > 4);
  print(await s.join(','));
}"#,
        ["6,8,10"]
    };

    async_star_yield_after_nested_await => {
        r#"Future<int> deep() async {
  return await Future.value(12);
}
Stream<int> gen() async* {
  yield await deep();
}
Future<void> main() async {
  print(await gen().first);
}"#,
        ["12"]
    };

    async_star_infinite_with_take_is_safe => {
        r#"Stream<int> ids() async* {
  var n = 0;
  while (true) { yield n; n++; }
}
Future<void> main() async {
  print(await ids().take(3).join(','));
}"#,
        ["0,1,2"]
    };

}
