//! Future combinators — wait, any, then, catchError, whenComplete.

dart_cases! {
    future_wait_all_success_returns_ordered_values => {
        r#"Future<void> main() async {
  var results = await Future.wait([Future.value(1), Future.value(2), Future.value(3)]);
  print(results.join(','));
}"#,
        ["1,2,3"]
    };

    future_wait_empty_list_returns_empty => {
        r#"Future<void> main() async {
  var results = await Future.wait<int>([]);
  print(results.length);
}"#,
        ["0"]
    };

    future_wait_single_future_returns_one_element => {
        r#"Future<void> main() async {
  var results = await Future.wait([Future.value('solo')]);
  print(results[0]);
}"#,
        ["solo"]
    };

    future_wait_one_fails_propagates_error => {
        r#"Future<void> main() async {
  try {
    await Future.wait([Future.value(1), Future<int>.error('bad')]);
    print('ok');
  } catch (e) {
    print('err:$e');
  }
}"#,
        ["err:bad"]
    };

    future_wait_with_async_functions => {
        r#"Future<int> doubleIt(int n) async => n * 2;
Future<void> main() async {
  var results = await Future.wait([doubleIt(2), doubleIt(3)]);
  print(results.join(','));
}"#,
        ["4,6"]
    };

    future_any_first_completing_value_wins => {
        r#"Future<void> main() async {
  var v = await Future.any([Future.value(10), Future.value(20)]);
  print(v);
}"#,
        ["10"]
    };

    future_any_single_future_returns_its_value => {
        r#"Future<void> main() async {
  var v = await Future.any([Future.value(7)]);
  print(v);
}"#,
        ["7"]
    };

    future_any_with_error_future_throws => {
        r#"Future<void> main() async {
  try {
    await Future.any([Future<int>.error('nope')]);
    print('ok');
  } catch (e) {
    print('err:$e');
  }
}"#,
        ["err:nope"]
    };

    future_sync_completes_with_returned_value => {
        r#"Future<void> main() async {
  var v = await Future.sync(() => 42);
  print(v);
}"#,
        ["42"]
    };

    future_sync_computes_string_result => {
        r#"Future<void> main() async {
  var s = await Future.sync(() => 'computed');
  print(s);
}"#,
        ["computed"]
    };

    future_sync_throws_becomes_error_future => {
        r#"Future<void> main() async {
  try {
    await Future.sync(() => throw 'sync-fail');
  } catch (e) {
    print('err:$e');
  }
}"#,
        ["err:sync-fail"]
    };

    future_error_completes_with_error_value => {
        r#"Future<void> main() async {
  try {
    await Future<int>.error('boom');
  } catch (e) {
    print('err:$e');
  }
}"#,
        ["err:boom"]
    };

    future_value_int_awaited_directly => {
        r#"Future<void> main() async {
  var v = await Future.value(99);
  print(v);
}"#,
        ["99"]
    };

    future_value_string_awaited_directly => {
        r#"Future<void> main() async {
  var s = await Future.value('ready');
  print(s);
}"#,
        ["ready"]
    };

    future_value_bool_awaited_directly => {
        r#"Future<void> main() async {
  var b = await Future.value(false);
  print(b);
}"#,
        ["false"]
    };

    future_value_null_awaited => {
        r#"Future<void> main() async {
  int? v = await Future<int?>.value(null);
  print(v == null);
}"#,
        ["true"]
    };

    future_then_transforms_completed_value => {
        r#"Future<void> main() async {
  var v = await Future.value(3).then((x) => x + 4);
  print(v);
}"#,
        ["7"]
    };

    future_then_chain_three_transformations => {
        r#"Future<void> main() async {
  var v = await Future.value(2)
      .then((x) => x + 1)
      .then((x) => x * 3)
      .then((x) => x - 2);
  print(v);
}"#,
        ["7"]
    };

    future_then_returns_new_future_type => {
        r#"Future<void> main() async {
  var s = await Future.value(5).then((x) => 'n$x');
  print(s);
}"#,
        ["n5"]
    };

    future_catch_error_returns_recovery_value => {
        r#"Future<void> main() async {
  var v = await Future<int>.error('x').catchError((e) => 0);
  print(v);
}"#,
        ["0"]
    };

    future_catch_error_test_only_handles_matching => {
        r#"Future<void> main() async {
  try {
    await Future<int>.error('fail').catchError((e) => 1, test: (e) => e == 'other');
  } catch (e) {
    print('err:$e');
  }
}"#,
        ["err:fail"]
    };

    future_catch_error_test_handles_matching_error => {
        r#"Future<void> main() async {
  var v = await Future<int>.error('recover').catchError((e) => 9, test: (e) => e == 'recover');
  print(v);
}"#,
        ["9"]
    };

    future_when_complete_runs_after_success => {
        r#"Future<void> main() async {
  var log = <String>[];
  await Future.value(1).whenComplete(() => log.add('done'));
  print(log.join(','));
}"#,
        ["done"]
    };

    future_when_complete_runs_after_error => {
        r#"Future<void> main() async {
  var log = <String>[];
  try {
    await Future<int>.error('e').whenComplete(() => log.add('cleanup'));
  } catch (_) {}
  print(log.join(','));
}"#,
        ["cleanup"]
    };

    future_then_catch_error_when_complete_execution_order => {
        r#"Future<void> main() async {
  var log = <String>[];
  await Future.value(1)
      .then((v) {
        log.add('then');
        return v;
      })
      .catchError((e) {
        log.add('catch');
        return 0;
      })
      .whenComplete(() => log.add('complete'));
  print(log.join(','));
}"#,
        ["then,complete"]
    };

    future_error_then_skipped_catch_error_runs => {
        r#"Future<void> main() async {
  var log = <String>[];
  var v = await Future<int>.error('oops')
      .then((x) {
        log.add('then');
        return x;
      })
      .catchError((e) {
        log.add('catch');
        return 1;
      });
  print('$v|${log.join(',')}');
}"#,
        ["1|catch"]
    };

    future_when_complete_does_not_swallow_error => {
        r#"Future<void> main() async {
  var log = <String>[];
  try {
    await Future<int>.error('fail').whenComplete(() => log.add('wc'));
  } catch (e) {
    log.add('thrown');
  }
  print(log.join(','));
}"#,
        ["wc,thrown"]
    };

    async_main_sequential_two_awaits => {
        r#"Future<int> stepA() async => 10;
Future<int> stepB(int n) async => n + 5;
Future<void> main() async {
  var a = await stepA();
  var b = await stepB(a);
  print(b);
}"#,
        ["15"]
    };

    async_main_three_awaits_accumulates_sum => {
        r#"Future<int> one() async => 1;
Future<int> two() async => 2;
Future<int> three() async => 3;
Future<void> main() async {
  var total = await one() + await two() + await three();
  print(total);
}"#,
        ["6"]
    };

    async_main_await_in_for_loop_builds_list => {
        r#"Future<int> id(int n) async => n;
Future<void> main() async {
  var out = <int>[];
  for (var i = 1; i <= 3; i++) {
    out.add(await id(i));
  }
  print(out.join(','));
}"#,
        ["1,2,3"]
    };

    async_main_await_inside_if_branch => {
        r#"Future<String> pick(bool flag) async => flag ? 'yes' : 'no';
Future<void> main() async {
  var s = '';
  if (true) {
    s = await pick(true);
  } else {
    s = await pick(false);
  }
  print(s);
}"#,
        ["yes"]
    };

    async_main_nested_await_calls => {
        r#"Future<int> inner() async => 4;
Future<int> outer() async => await inner() + 1;
Future<void> main() async {
  print(await outer());
}"#,
        ["5"]
    };

    future_wait_string_results_joined => {
        r#"Future<void> main() async {
  var results = await Future.wait([Future.value('a'), Future.value('b')]);
  print(results.join(''));
}"#,
        ["ab"]
    };

    future_wait_mixed_success_values_by_index => {
        r#"Future<void> main() async {
  var results = await Future.wait([Future.value(1), Future.value('two')]);
  print('${results[0]}|${results[1]}');
}"#,
        ["1|two"]
    };

    future_sync_zero_value => {
        r#"Future<void> main() async {
  print(await Future.sync(() => 0));
}"#,
        ["0"]
    };

    future_sync_negative_value => {
        r#"Future<void> main() async {
  print(await Future.sync(() => -5));
}"#,
        ["-5"]
    };

    future_then_on_future_sync_result => {
        r#"Future<void> main() async {
  var v = await Future.sync(() => 6).then((x) => x * 2);
  print(v);
}"#,
        ["12"]
    };

    future_catch_error_on_future_sync_throw => {
        r#"Future<void> main() async {
  var v = await Future.sync<int>(() => throw 'sync').catchError((e) => 3);
  print(v);
}"#,
        ["3"]
    };

    future_value_double_awaited => {
        r#"Future<void> main() async {
  print(await Future.value(2.5));
}"#,
        ["2.5"]
    };

    future_then_with_async_callback => {
        r#"Future<int> bump(int n) async => n + 1;
Future<void> main() async {
  var v = await Future.value(8).then((x) => bump(x));
  print(v);
}"#,
        ["9"]
    };

    future_wait_two_async_computed_values => {
        r#"Future<int> square(int n) async => n * n;
Future<void> main() async {
  var results = await Future.wait([square(2), square(3)]);
  print(results.join(','));
}"#,
        ["4,9"]
    };

    future_any_among_two_values_returns_one => {
        r#"Future<void> main() async {
  var v = await Future.any([Future.value('first'), Future.value('second')]);
  print(v == 'first' || v == 'second');
}"#,
        ["true"]
    };

    catch_error_chained_with_second_handler => {
        r#"Future<void> main() async {
  var v = await Future<int>.error('a')
      .catchError((e) => throw 'b')
      .catchError((e) => 7);
  print(v);
}"#,
        ["7"]
    };

    when_complete_followed_by_then_on_success => {
        r#"Future<void> main() async {
  var log = <String>[];
  var v = await Future.value(2)
      .whenComplete(() => log.add('wc'))
      .then((x) {
        log.add('then');
        return x + 1;
      });
  print('$v|${log.join(',')}');
}"#,
        ["3|wc,then"]
    };

    async_main_early_return_after_single_await => {
        r#"Future<int> load() async => 50;
Future<int> run() async {
  var v = await load();
  return v;
}
Future<void> main() async {
  print(await run());
}"#,
        ["50"]
    };

    future_wait_result_length_matches_input => {
        r#"Future<void> main() async {
  var results = await Future.wait([Future.value(1), Future.value(2), Future.value(3), Future.value(4)]);
  print(results.length);
}"#,
        ["4"]
    };

    future_error_string_type_preserved_in_catch => {
        r#"Future<void> main() async {
  Object? caught;
  try {
    await Future<String>.error('typed');
  } catch (e) {
    caught = e;
  }
  print(caught);
}"#,
        ["typed"]
    };

    async_main_await_expression_in_addition => {
        r#"Future<int> five() async => 5;
Future<void> main() async {
  print(2 + await five());
}"#,
        ["7"]
    };

    future_then_identity_returns_same_value => {
        r#"Future<void> main() async {
  var v = await Future.value(11).then((x) => x);
  print(v);
}"#,
        ["11"]
    };

    future_sync_side_effect_runs_before_await => {
        r#"Future<void> main() async {
  var log = <String>[];
  var v = await Future.sync(() {
    log.add('work');
    return 1;
  });
  log.add('after');
  print('${v}|${log.join(',')}');
}"#,
        ["1|work,after"]
    };

    async_main_multiple_local_async_functions => {
        r#"Future<void> main() async {
  Future<int> a() async => 1;
  Future<int> b() async => 2;
  print(await a() + await b());
}"#,
        ["3"]
    };

    future_wait_all_bool_results => {
        r#"Future<void> main() async {
  var results = await Future.wait([Future.value(true), Future.value(false), Future.value(true)]);
  print(results.where((b) => b).length);
}"#,
        ["2"]
    };

    future_catch_error_with_on_error_callback_style => {
        r#"Future<void> main() async {
  var v = await Future<int>.error('msg').catchError((e, st) => 42);
  print(v);
}"#,
        ["42"]
    };

    async_main_while_loop_with_await => {
        r#"Future<int> tick(int n) async => n;
Future<void> main() async {
  var i = 0;
  var sum = 0;
  while (i < 3) {
    sum = sum + await tick(i + 1);
    i++;
  }
  print(sum);
}"#,
        ["6"]
    };

    future_value_then_when_complete_order_on_error_path => {
        r#"Future<void> main() async {
  var log = <String>[];
  try {
    await Future<int>.error('e')
        .then((x) {
          log.add('then');
          return x;
        })
        .whenComplete(() => log.add('wc'));
  } catch (_) {
    log.add('catch');
  }
  print(log.join(','));
}"#,
        ["wc,catch"]
    };

    future_sync_bool_result => {
        r#"Future<void> main() async {
  print(await Future.sync(() => true));
}"#,
        ["true"]
    };

    async_main_await_nullable_with_coalesce => {
        r#"Future<int?> maybe() async => null;
Future<void> main() async {
  print(await maybe() ?? 8);
}"#,
        ["8"]
    };
}
