//! Streams — listen, transform, collect.

dart_cases! {
    stream_from_iterable_emits_all_elements => {
        r#"Future<void> main() async {
  var out = <int>[];
  await for (var v in Stream.fromIterable([1, 2, 3])) { out.add(v); }
  print(out.join(','));
}"#,
        ["1,2,3"]
    };

    stream_from_iterable_empty_yields_nothing => {
        r#"Future<void> main() async {
  var out = <int>[];
  await for (var v in Stream.fromIterable(<int>[])) { out.add(v); }
  print(out.length);
}"#,
        ["0"]
    };

    stream_empty_has_no_events => {
        r#"Future<void> main() async {
  var count = 0;
  await for (var _ in Stream<int>.empty()) { count++; }
  print(count);
}"#,
        ["0"]
    };

    stream_value_emits_single_element => {
        r#"Future<void> main() async {
  var out = <int>[];
  await for (var v in Stream.value(42)) { out.add(v); }
  print(out.join(','));
}"#,
        ["42"]
    };

    stream_value_string_emits_once => {
        r#"Future<void> main() async {
  var out = <String>[];
  await for (var v in Stream.value('solo')) { out.add(v); }
  print(out.join('|'));
}"#,
        ["solo"]
    };

    stream_map_doubles_each_int => {
        r#"Future<void> main() async {
  var out = <int>[];
  var s = Stream.fromIterable([1, 2, 3]).map((x) => x * 2);
  await for (var v in s) { out.add(v); }
  print(out.join(','));
}"#,
        ["2,4,6"]
    };

    stream_map_converts_int_to_string => {
        r#"Future<void> main() async {
  var out = <String>[];
  await for (var v in Stream.fromIterable([1, 2]).map((x) => 'n$x')) { out.add(v); }
  print(out.join(','));
}"#,
        ["n1,n2"]
    };

    stream_where_keeps_only_evens => {
        r#"Future<void> main() async {
  var out = <int>[];
  await for (var v in Stream.fromIterable([1, 2, 3, 4]).where((x) => x % 2 == 0)) { out.add(v); }
  print(out.join(','));
}"#,
        ["2,4"]
    };

    stream_where_rejects_all_when_predicate_false => {
        r#"Future<void> main() async {
  var count = 0;
  await for (var _ in Stream.fromIterable([1, 2, 3]).where((x) => x > 10)) { count++; }
  print(count);
}"#,
        ["0"]
    };

    stream_take_limits_to_first_n => {
        r#"Future<void> main() async {
  var out = <int>[];
  await for (var v in Stream.fromIterable([10, 20, 30, 40]).take(2)) { out.add(v); }
  print(out.join(','));
}"#,
        ["10,20"]
    };

    stream_take_zero_emits_nothing => {
        r#"Future<void> main() async {
  var count = 0;
  await for (var _ in Stream.fromIterable([1, 2, 3]).take(0)) { count++; }
  print(count);
}"#,
        ["0"]
    };

    stream_skip_ignores_first_elements => {
        r#"Future<void> main() async {
  var out = <int>[];
  await for (var v in Stream.fromIterable([1, 2, 3, 4]).skip(2)) { out.add(v); }
  print(out.join(','));
}"#,
        ["3,4"]
    };

    stream_skip_beyond_length_yields_empty => {
        r#"Future<void> main() async {
  var count = 0;
  await for (var _ in Stream.fromIterable([1, 2]).skip(5)) { count++; }
  print(count);
}"#,
        ["0"]
    };

    stream_expand_flattens_inner_lists => {
        r#"Future<void> main() async {
  var out = <int>[];
  var nested = Stream.fromIterable([[1, 2], [3], [4, 5]]);
  await for (var v in nested.expand((part) => part)) { out.add(v); }
  print(out.join(','));
}"#,
        ["1,2,3,4,5"]
    };

    stream_expand_splits_strings_to_chars => {
        r#"Future<void> main() async {
  var out = <String>[];
  await for (var c in Stream.fromIterable(['ab', 'c']).expand((w) => w.split(''))) { out.add(c); }
  print(out.join(''));
}"#,
        ["abc"]
    };

    stream_async_map_applies_async_transform => {
        r#"Future<void> main() async {
  var out = <int>[];
  await for (var v in Stream.fromIterable([1, 2, 3]).asyncMap((x) async => x + 10)) { out.add(v); }
  print(out.join(','));
}"#,
        ["11,12,13"]
    };

    stream_listen_on_data_collects_values => {
        r#"class DataSink {
  var values = <int>[];
}
Future<void> main() async {
  var sink = DataSink();
  var sub = Stream.fromIterable([4, 5, 6]).listen((v) => sink.values.add(v));
  await sub.asFuture();
  print(sink.values.join(','));
}"#,
        ["4,5,6"]
    };

    stream_listen_on_done_increments_counter => {
        r#"class DoneTracker {
  var doneCount = 0;
}
Future<void> main() async {
  var t = DoneTracker();
  var sub = Stream.fromIterable([1]).listen(
    (_) {},
    onDone: () => t.doneCount++,
  );
  await sub.asFuture();
  print(t.doneCount);
}"#,
        ["1"]
    };

    stream_listen_on_error_increments_error_counter => {
        r#"class ErrorTracker {
  var errorCount = 0;
  String? lastError;
}
Future<void> main() async {
  var t = ErrorTracker();
  var sub = Stream<int>.error('boom').listen(
    (_) {},
    onError: (e) {
      t.errorCount++;
      t.lastError = '$e';
    },
  );
  try {
    await sub.asFuture();
  } catch (_) {}
  print('${t.errorCount}|${t.lastError}');
}"#,
        ["1|boom"]
    };

    stream_listen_data_and_done_together => {
        r#"class StreamStats {
  var dataCount = 0;
  var done = false;
}
Future<void> main() async {
  var s = StreamStats();
  var sub = Stream.fromIterable([1, 2, 3]).listen(
    (_) => s.dataCount++,
    onDone: () => s.done = true,
  );
  await sub.asFuture();
  print('${s.dataCount}|${s.done}');
}"#,
        ["3|true"]
    };

    stream_to_list_collects_in_order => {
        r#"Future<void> main() async {
  var list = await Stream.fromIterable(['a', 'b', 'c']).toList();
  print(list.join(','));
}"#,
        ["a,b,c"]
    };

    stream_first_returns_initial_element => {
        r#"Future<void> main() async {
  var v = await Stream.fromIterable([7, 8, 9]).first;
  print(v);
}"#,
        ["7"]
    };

    stream_last_returns_final_element => {
        r#"Future<void> main() async {
  var v = await Stream.fromIterable([7, 8, 9]).last;
  print(v);
}"#,
        ["9"]
    };

    stream_length_counts_all_events => {
        r#"Future<void> main() async {
  var n = await Stream.fromIterable([1, 2, 3, 4]).length;
  print(n);
}"#,
        ["4"]
    };

    stream_is_empty_true_for_empty_stream => {
        r#"Future<void> main() async {
  var empty = await Stream<int>.empty().isEmpty;
  print(empty);
}"#,
        ["true"]
    };

    stream_is_empty_false_for_nonempty_stream => {
        r#"Future<void> main() async {
  var empty = await Stream.value(1).isEmpty;
  print(empty);
}"#,
        ["false"]
    };

    stream_distinct_removes_consecutive_duplicates => {
        r#"Future<void> main() async {
  var out = <int>[];
  await for (var v in Stream.fromIterable([1, 1, 2, 2, 2, 3]).distinct()) { out.add(v); }
  print(out.join(','));
}"#,
        ["1,2,3"]
    };

    stream_distinct_keeps_non_adjacent_duplicates => {
        r#"Future<void> main() async {
  var out = <int>[];
  await for (var v in Stream.fromIterable([1, 2, 1, 2, 1]).distinct()) { out.add(v); }
  print(out.join(','));
}"#,
        ["1,2,1,2,1"]
    };

    stream_distinct_with_equals_callback => {
        r#"Future<void> main() async {
  var out = <String>[];
  await for (var v in Stream.fromIterable(['aa', 'AA', 'bb', 'BB']).distinct((a, b) => a.toLowerCase() == b.toLowerCase())) {
    out.add(v);
  }
  print(out.join(','));
}"#,
        ["aa,bb"]
    };

    stream_distinct_on_empty_stream => {
        r#"Future<void> main() async {
  var list = await Stream<int>.empty().distinct().toList();
  print(list.length);
}"#,
        ["0"]
    };

    stream_chain_map_then_where => {
        r#"Future<void> main() async {
  var out = <int>[];
  var s = Stream.fromIterable([1, 2, 3, 4, 5]).map((x) => x * 2).where((x) => x > 4);
  await for (var v in s) { out.add(v); }
  print(out.join(','));
}"#,
        ["6,8,10"]
    };

    stream_chain_skip_then_take => {
        r#"Future<void> main() async {
  var out = <int>[];
  await for (var v in Stream.fromIterable([1, 2, 3, 4, 5, 6]).skip(2).take(2)) { out.add(v); }
  print(out.join(','));
}"#,
        ["3,4"]
    };

    stream_single_subscription_listen_once_succeeds => {
        r#"class OnceListener {
  var count = 0;
}
Future<void> main() async {
  var o = OnceListener();
  var sub = Stream.value(99).listen((v) => o.count = v);
  await sub.asFuture();
  print(o.count);
}"#,
        ["99"]
    };

    stream_for_each_runs_callback_per_event => {
        r#"Future<void> main() async {
  var sum = 0;
  await Stream.fromIterable([1, 2, 3]).forEach((v) => sum = sum + v);
  print(sum);
}"#,
        ["6"]
    };

    stream_contains_finds_present_value => {
        r#"Future<void> main() async {
  var found = await Stream.fromIterable([2, 4, 6]).contains(4);
  print(found);
}"#,
        ["true"]
    };

    stream_contains_returns_false_for_missing => {
        r#"Future<void> main() async {
  var found = await Stream.fromIterable([2, 4, 6]).contains(5);
  print(found);
}"#,
        ["false"]
    };

    stream_every_all_match_predicate => {
        r#"Future<void> main() async {
  var ok = await Stream.fromIterable([2, 4, 6]).every((x) => x % 2 == 0);
  print(ok);
}"#,
        ["true"]
    };

    stream_every_fails_when_one_does_not_match => {
        r#"Future<void> main() async {
  var ok = await Stream.fromIterable([2, 3, 4]).every((x) => x % 2 == 0);
  print(ok);
}"#,
        ["false"]
    };

    stream_reduce_combines_all_elements => {
        r#"Future<void> main() async {
  var product = await Stream.fromIterable([2, 3, 4]).reduce((a, b) => a * b);
  print(product);
}"#,
        ["24"]
    };

    stream_fold_with_seed_accumulates => {
        r#"Future<void> main() async {
  var total = await Stream.fromIterable([1, 2, 3]).fold(10, (acc, v) => acc + v);
  print(total);
}"#,
        ["16"]
    };

    stream_element_at_returns_indexed_value => {
        r#"Future<void> main() async {
  var v = await Stream.fromIterable(['p', 'q', 'r']).elementAt(1);
  print(v);
}"#,
        ["q"]
    };

    stream_take_while_stops_at_first_false => {
        r#"Future<void> main() async {
  var out = <int>[];
  await for (var v in Stream.fromIterable([1, 2, 3, 1, 2]).takeWhile((x) => x < 3)) { out.add(v); }
  print(out.join(','));
}"#,
        ["1,2"]
    };

    stream_skip_while_skips_matching_prefix => {
        r#"Future<void> main() async {
  var out = <int>[];
  await for (var v in Stream.fromIterable([0, 0, 1, 2, 0]).skipWhile((x) => x == 0)) { out.add(v); }
  print(out.join(','));
}"#,
        ["1,2,0"]
    };

    stream_await_for_break_exits_early => {
        r#"Future<void> main() async {
  var out = <int>[];
  await for (var v in Stream.fromIterable([1, 2, 3, 4, 5])) {
    if (v == 3) break;
    out.add(v);
  }
  print(out.join(','));
}"#,
        ["1,2"]
    };

    stream_await_for_continue_skips_iteration => {
        r#"Future<void> main() async {
  var out = <int>[];
  await for (var v in Stream.fromIterable([1, 2, 3, 4])) {
    if (v % 2 == 0) continue;
    out.add(v);
  }
  print(out.join(','));
}"#,
        ["1,3"]
    };

    stream_value_null_element_emitted => {
        r#"Future<void> main() async {
  var out = <int?>[];
  await for (var v in Stream<int?>.value(null)) { out.add(v); }
  print(out.length);
}"#,
        ["1"]
    };

    stream_from_iterable_strings_joined => {
        r#"Future<void> main() async {
  var list = await Stream.fromIterable(['x', 'y', 'z']).toList();
  print(list.join(''));
}"#,
        ["xyz"]
    };

    stream_map_identity_preserves_values => {
        r#"Future<void> main() async {
  var out = <int>[];
  await for (var v in Stream.fromIterable([5, 6, 7]).map((x) => x)) { out.add(v); }
  print(out.join(','));
}"#,
        ["5,6,7"]
    };

    stream_where_on_strings_filters_length => {
        r#"Future<void> main() async {
  var out = <String>[];
  await for (var v in Stream.fromIterable(['a', 'bb', 'ccc']).where((s) => s.length == 2)) { out.add(v); }
  print(out.join(','));
}"#,
        ["bb"]
    };

    stream_expand_empty_inner_yields_nothing => {
        r#"Future<void> main() async {
  var count = 0;
  await for (var _ in Stream.fromIterable(<List<int>>[[]]).expand((p) => p)) { count++; }
  print(count);
}"#,
        ["0"]
    };

    stream_to_list_after_map_transform => {
        r#"Future<void> main() async {
  var list = await Stream.fromIterable([1, 2, 3]).map((x) => x + 1).toList();
  print(list.join(','));
}"#,
        ["2,3,4"]
    };

    stream_first_on_value_stream => {
        r#"Future<void> main() async {
  var v = await Stream.value('only').first;
  print(v);
}"#,
        ["only"]
    };

    stream_last_on_value_stream => {
        r#"Future<void> main() async {
  var v = await Stream.value('only').last;
  print(v);
}"#,
        ["only"]
    };

    stream_length_of_value_stream_is_one => {
        r#"Future<void> main() async {
  var n = await Stream.value(0).length;
  print(n);
}"#,
        ["1"]
    };

    stream_listen_cancel_stops_before_done_flag => {
        r#"class CancelTracker {
  var dataCount = 0;
  var cancelled = false;
}
Future<void> main() async {
  var t = CancelTracker();
  late StreamSubscription<int> sub;
  sub = Stream.fromIterable([1, 2, 3, 4]).listen((v) {
    t.dataCount++;
    if (v == 2) {
      sub.cancel();
      t.cancelled = true;
    }
  });
  try {
    await sub.asFuture();
  } catch (_) {}
  print('${t.dataCount}|${t.cancelled}');
}"#,
        ["2|true"]
    };

    stream_pipeline_map_where_take => {
        r#"Future<void> main() async {
  var out = <int>[];
  var s = Stream.fromIterable([1, 2, 3, 4, 5, 6])
      .map((x) => x * 3)
      .where((x) => x > 6)
      .take(2);
  await for (var v in s) { out.add(v); }
  print(out.join(','));
}"#,
        ["9,12"]
    };
}
