//! Dart sync* generators — yield, yield*, lazy iteration, Iterable consumption.

dart_cases! {
    sync_star_yields_single_int_via_for_in => {
        r#"Iterable<int> one() sync* { yield 42; }
void main() {
  for (var v in one()) { print(v); }
}"#,
        ["42"]
    };

    sync_star_yields_two_values_in_order => {
        r#"Iterable<int> pair() sync* { yield 1; yield 2; }
void main() {
  var out = <int>[];
  for (var v in pair()) { out.add(v); }
  print(out.join(','));
}"#,
        ["1,2"]
    };

    sync_star_yields_three_values_materialized => {
        r#"Iterable<int> triple() sync* { yield 10; yield 20; yield 30; }
void main() {
  print(triple().toList().join('|'));
}"#,
        ["10|20|30"]
    };

    sync_star_yields_string_elements => {
        r#"Iterable<String> words() sync* { yield 'a'; yield 'b'; yield 'c'; }
void main() {
  print(words().join(''));
}"#,
        ["abc"]
    };

    sync_star_yields_bool_values => {
        r#"Iterable<bool> flags() sync* { yield true; yield false; yield true; }
void main() {
  var out = <String>[];
  for (var f in flags()) { out.add('$f'); }
  print(out.join(','));
}"#,
        ["true,false,true"]
    };

    sync_star_yields_doubles => {
        r#"Iterable<double> halves() sync* { yield 0.5; yield 1.5; }
void main() {
  var sum = 0.0;
  for (var v in halves()) { sum += v; }
  print(sum);
}"#,
        ["2"]
    };

    sync_star_empty_generator_produces_no_elements => {
        r#"Iterable<int> empty() sync* {}
void main() {
  print(empty().length);
}"#,
        ["0"]
    };

    sync_star_single_yield_then_completes => {
        r#"Iterable<int> solo() sync* { yield 7; }
void main() {
  print(solo().first);
}"#,
        ["7"]
    };

    sync_star_range_via_for_loop => {
        r#"Iterable<int> range(int n) sync* {
  for (var i = 0; i < n; i++) { yield i; }
}
void main() {
  print(range(4).join(','));
}"#,
        ["0,1,2,3"]
    };

    sync_star_while_loop_yields_countdown => {
        r#"Iterable<int> countdown(int n) sync* {
  while (n > 0) { yield n; n--; }
}
void main() {
  print(countdown(3).join(','));
}"#,
        ["3,2,1"]
    };

    sync_star_evens_via_step_loop => {
        r#"Iterable<int> evens(int limit) sync* {
  for (var i = 0; i < limit; i += 2) { yield i; }
}
void main() {
  print(evens(8).join(','));
}"#,
        ["0,2,4,6"]
    };

    sync_star_odds_via_continue_skip_evens => {
        r#"Iterable<int> odds(int n) sync* {
  for (var i = 0; i < n; i++) {
    if (i % 2 == 0) continue;
    yield i;
  }
}
void main() {
  print(odds(6).join(','));
}"#,
        ["1,3,5"]
    };

    sync_star_break_inside_loop_limits_yields => {
        r#"Iterable<int> capped() sync* {
  for (var i = 0; i < 10; i++) {
    if (i == 4) break;
    yield i;
  }
}
void main() {
  print(capped().join(','));
}"#,
        ["0,1,2,3"]
    };

    sync_star_if_branch_selects_yield_path => {
        r#"Iterable<int> branch(bool pickA) sync* {
  if (pickA) { yield 1; } else { yield 2; }
}
void main() {
  print(branch(true).first);
}"#,
        ["1"]
    };

    sync_star_if_else_yields_alternate_value => {
        r#"Iterable<int> branch(bool pickA) sync* {
  if (pickA) { yield 1; } else { yield 2; }
}
void main() {
  print(branch(false).first);
}"#,
        ["2"]
    };

    sync_star_do_while_yields_at_least_once => {
        r#"Iterable<int> once() sync* {
  var i = 0;
  do { yield i; i++; } while (i < 1);
}
void main() {
  print(once().join(','));
}"#,
        ["0"]
    };

    sync_star_yield_star_delegates_to_inner_sync_star => {
        r#"Iterable<int> inner() sync* { yield 1; yield 2; }
Iterable<int> outer() sync* { yield* inner(); }
void main() {
  print(outer().join(','));
}"#,
        ["1,2"]
    };

    sync_star_yield_star_then_extra_yield => {
        r#"Iterable<int> inner() sync* { yield 1; yield 2; }
Iterable<int> outer() sync* { yield* inner(); yield 3; }
void main() {
  print(outer().join(','));
}"#,
        ["1,2,3"]
    };

    sync_star_yield_star_from_list_literal => {
        r#"Iterable<int> fromList() sync* { yield* [5, 6, 7]; }
void main() {
  print(fromList().join(','));
}"#,
        ["5,6,7"]
    };

    sync_star_yield_star_empty_list_yields_nothing => {
        r#"Iterable<int> fromEmpty() sync* { yield* <int>[]; yield 99; }
void main() {
  print(fromEmpty().join(','));
}"#,
        ["99"]
    };

    sync_star_multiple_yield_star_in_sequence => {
        r#"Iterable<int> a() sync* { yield 1; }
Iterable<int> b() sync* { yield 2; }
Iterable<int> both() sync* { yield* a(); yield* b(); yield 3; }
void main() {
  print(both().join(','));
}"#,
        ["1,2,3"]
    };

    sync_star_yield_star_from_string_chars => {
        r#"Iterable<String> chars() sync* { yield* 'hi'; }
void main() {
  print(chars().join(''));
}"#,
        ["hi"]
    };

    sync_star_nested_yield_star_chain => {
        r#"Iterable<int> a() sync* { yield 1; }
Iterable<int> b() sync* { yield* a(); yield 2; }
Iterable<int> c() sync* { yield* b(); yield 3; }
void main() {
  print(c().join(','));
}"#,
        ["1,2,3"]
    };

    sync_star_yield_star_preserves_order_with_prefix => {
        r#"Iterable<int> inner() sync* { yield 2; yield 3; }
Iterable<int> outer() sync* { yield 1; yield* inner(); yield 4; }
void main() {
  print(outer().join(','));
}"#,
        ["1,2,3,4"]
    };

    sync_star_nested_generator_function_call => {
        r#"Iterable<int> inner() sync* { yield 10; }
Iterable<int> outer() sync* { yield* inner(); yield 20; }
void main() {
  print(outer().length);
}"#,
        ["2"]
    };

    sync_star_recursive_countdown_generator => {
        r#"Iterable<int> down(int n) sync* {
  if (n <= 0) return;
  yield n;
  yield* down(n - 1);
}
void main() {
  print(down(3).join(','));
}"#,
        ["3,2,1"]
    };

    sync_star_recursive_yields_prefix_then_suffix => {
        r#"Iterable<int> walk(int n) sync* {
  if (n == 0) return;
  yield n;
  yield* walk(n - 1);
}
void main() {
  print(walk(2).join(','));
}"#,
        ["2,1"]
    };

    sync_star_generator_with_parameters => {
        r#"Iterable<int> repeat(int value, int times) sync* {
  for (var i = 0; i < times; i++) { yield value; }
}
void main() {
  print(repeat(9, 3).join(','));
}"#,
        ["9,9,9"]
    };

    sync_star_closure_captures_outer_variable => {
        r#"Iterable<int> make() sync* {
  var base = 5;
  yield base;
  base = base + 1;
  yield base;
}
void main() {
  print(make().join(','));
}"#,
        ["5,6"]
    };

    sync_star_local_state_persists_between_yields => {
        r#"Iterable<int> counter() sync* {
  var n = 0;
  yield n;
  n = n + 1;
  yield n;
  n = n + 1;
  yield n;
}
void main() {
  print(counter().join(','));
}"#,
        ["0,1,2"]
    };

    sync_star_tree_preorder_traversal => {
        r#"Iterable<int> preorder(List<int> nodes) sync* {
  if (nodes.isEmpty) return;
  yield nodes[0];
  if (nodes.length > 1) { yield* preorder(nodes.sublist(1)); }
}
void main() {
  print(preorder([1, 2, 3]).join(','));
}"#,
        ["1,2,3"]
    };

    sync_star_depth_first_yields_leaves => {
        r#"Iterable<String> leaves() sync* {
  yield 'a';
  yield 'b';
  yield* leavesHelper();
}
Iterable<String> leavesHelper() sync* { yield 'c'; }
void main() {
  print(leaves().join(''));
}"#,
        ["abc"]
    };

    sync_star_fibonacci_first_six_via_take => {
        r#"Iterable<int> fib() sync* {
  var a = 0, b = 1;
  while (true) { yield a; var c = a + b; a = b; b = c; }
}
void main() {
  print(fib().take(6).join(','));
}"#,
        ["0,1,1,2,3,5"]
    };

    sync_star_fibonacci_first_ten_sum => {
        r#"Iterable<int> fib() sync* {
  var a = 0, b = 1;
  while (true) { yield a; var c = a + b; a = b; b = c; }
}
void main() {
  var sum = 0;
  for (var v in fib().take(10)) { sum += v; }
  print(sum);
}"#,
        ["88"]
    };

    sync_star_manual_take_stops_after_n_items => {
        r#"Iterable<int> naturals() sync* {
  var n = 1;
  while (true) { yield n; n++; }
}
void main() {
  var out = <int>[];
  var taken = 0;
  for (var v in naturals()) {
    out.add(v);
    taken++;
    if (taken == 3) break;
  }
  print(out.join(','));
}"#,
        ["1,2,3"]
    };

    sync_star_take_while_stops_at_threshold => {
        r#"Iterable<int> growing() sync* {
  for (var i = 1; i <= 10; i++) { yield i; }
}
void main() {
  print(growing().takeWhile((x) => x < 4).join(','));
}"#,
        ["1,2,3"]
    };

    sync_star_skip_then_take_slices_sequence => {
        r#"Iterable<int> seq() sync* {
  for (var i = 0; i < 6; i++) { yield i; }
}
void main() {
  print(seq().skip(2).take(2).join(','));
}"#,
        ["2,3"]
    };

    sync_star_infinite_generator_safe_with_take => {
        r#"Iterable<int> infinite() sync* {
  var n = 0;
  while (true) { yield n; n++; }
}
void main() {
  print(infinite().take(4).join(','));
}"#,
        ["0,1,2,3"]
    };

    sync_star_is_iterable_via_for_in => {
        r#"Iterable<int> gen() sync* { yield 4; yield 5; }
void main() {
  var sum = 0;
  for (var v in gen()) { sum += v; }
  print(sum);
}"#,
        ["9"]
    };

    sync_star_map_transforms_yields => {
        r#"Iterable<int> gen() sync* { yield 1; yield 2; yield 3; }
void main() {
  print(gen().map((x) => x * 10).join(','));
}"#,
        ["10,20,30"]
    };

    sync_star_where_filters_yields => {
        r#"Iterable<int> gen() sync* { for (var i = 1; i <= 5; i++) yield i; }
void main() {
  print(gen().where((x) => x % 2 == 1).join(','));
}"#,
        ["1,3,5"]
    };

    sync_star_expand_flattens_inner_iterables => {
        r#"Iterable<List<int>> gen() sync* {
  yield [1, 2];
  yield [3];
}
void main() {
  print(gen().expand((p) => p).join(','));
}"#,
        ["1,2,3"]
    };

    sync_star_followed_by_concatenates_tail => {
        r#"Iterable<int> gen() sync* { yield 1; yield 2; }
void main() {
  print(gen().followedBy([3, 4]).join(','));
}"#,
        ["1,2,3,4"]
    };

    sync_star_fold_reduces_all_yields => {
        r#"Iterable<int> gen() sync* { yield 1; yield 2; yield 3; }
void main() {
  print(gen().fold(0, (a, b) => a + b));
}"#,
        ["6"]
    };

    sync_star_reduce_multiplies_yields => {
        r#"Iterable<int> gen() sync* { yield 2; yield 3; yield 4; }
void main() {
  print(gen().reduce((a, b) => a * b));
}"#,
        ["24"]
    };

    sync_star_contains_finds_yielded_value => {
        r#"Iterable<int> gen() sync* { yield 11; yield 22; }
void main() {
  print(gen().contains(22));
}"#,
        ["true"]
    };

    sync_star_element_at_reads_nth_yield => {
        r#"Iterable<String> gen() sync* { yield 'x'; yield 'y'; yield 'z'; }
void main() {
  print(gen().elementAt(1));
}"#,
        ["y"]
    };

    sync_star_last_returns_final_yield => {
        r#"Iterable<int> gen() sync* { yield 100; yield 200; yield 300; }
void main() {
  print(gen().last);
}"#,
        ["300"]
    };

    sync_star_body_not_run_until_iteration_starts => {
        r#"Iterable<int> loud() sync* {
  print('side');
  yield 1;
}
void main() {
  var it = loud();
  print('before');
  print(it.first);
}"#,
        ["before", "side", "1"]
    };

    sync_star_yield_in_try_finally_still_yields => {
        r#"Iterable<int> guarded() sync* {
  try { yield 1; yield 2; } finally { print('fin'); }
}
void main() {
  print(guarded().join(','));
}"#,
        ["fin", "1,2"]
    };

    sync_star_multiple_generators_are_independent => {
        r#"Iterable<int> gen() sync* { yield 1; yield 2; }
void main() {
  var a = gen();
  var b = gen();
  print(a.first);
  print(b.first);
}"#,
        ["1", "1"]
    };

    sync_star_yield_after_awaiting_sync_work => {
        r#"Iterable<int> gen() sync* {
  var x = 2 + 3;
  yield x;
  yield x + 1;
}
void main() {
  print(gen().join(','));
}"#,
        ["5,6"]
    };

    sync_star_generator_used_in_list_comprehension_source => {
        r#"Iterable<int> gen() sync* { yield 1; yield 2; yield 3; }
void main() {
  var doubled = [for (var x in gen()) x * 2];
  print(doubled.join(','));
}"#,
        ["2,4,6"]
    };

    sync_star_to_set_deduplicates_yields => {
        r#"Iterable<int> gen() sync* { yield 1; yield 2; yield 2; yield 1; }
void main() {
  print(gen().toSet().length);
}"#,
        ["2"]
    };

    sync_star_any_detects_matching_yield => {
        r#"Iterable<int> gen() sync* { yield 2; yield 4; yield 6; }
void main() {
  print(gen().any((x) => x > 5));
}"#,
        ["true"]
    };

}
