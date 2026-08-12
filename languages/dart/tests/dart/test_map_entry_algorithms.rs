//! Map entry algorithms: entries.map, fold, removeWhere, update, putIfAbsent chains.

dart_cases! {
    entries_map_to_value_doubled_list => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  var doubled = m.entries.map((e) => e.value * 2).toList();
  print(doubled.join(','));
}"#,
        ["2,4,6"]
    };

    entries_map_to_key_value_strings => {
        r#"void main() {
  var m = {'x': 10, 'y': 20};
  var labels = m.entries.map((e) => '${e.key}=${e.value}').toList();
  print(labels.join('|'));
}"#,
        ["x=10|y=20"]
    };

    entries_map_builds_new_map_via_from_entries => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  var inverted = Map.fromEntries(
    m.entries.map((e) => MapEntry('${e.key}_x', e.value + 10)),
  );
  print(inverted['a_x']);
  print(inverted['b_x']);
  print(inverted.length);
}"#,
        ["11", "12", "2"]
    };

    entries_map_filters_then_collects_keys => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3, 'd': 4};
  var bigKeys = m.entries
      .where((e) => e.value > 2)
      .map((e) => e.key)
      .toList();
  print(bigKeys.join(','));
}"#,
        ["c,d"]
    };

    entries_map_to_map_with_transformed_values => {
        r#"void main() {
  var m = {'p': 1, 'q': 2};
  var squared = {for (var e in m.entries) e.key: e.value * e.value};
  print(squared['p']);
  print(squared['q']);
}"#,
        ["1", "4"]
    };

    entries_fold_sums_all_values => {
        r#"void main() {
  var m = {'a': 10, 'b': 20, 'c': 30};
  var total = m.entries.fold(0, (sum, e) => sum + e.value);
  print(total);
}"#,
        ["60"]
    };

    entries_fold_concatenates_keys => {
        r#"void main() {
  var m = {'one': 1, 'two': 2, 'three': 3};
  var keys = m.entries.fold('', (acc, e) => acc + e.key);
  print(keys);
}"#,
        ["onetwothree"]
    };

    entries_fold_builds_max_value => {
        r#"void main() {
  var m = {'a': 3, 'b': 9, 'c': 5};
  var maxVal = m.entries.fold(0, (best, e) => e.value > best ? e.value : best);
  print(maxVal);
}"#,
        ["9"]
    };

    entries_fold_counts_entries_above_threshold => {
        r#"void main() {
  var m = {'a': 1, 'b': 5, 'c': 8, 'd': 2};
  var count = m.entries.fold(0, (c, e) => e.value >= 5 ? c + 1 : c);
  print(count);
}"#,
        ["2"]
    };

    entries_fold_produces_comma_separated_pairs => {
        r#"void main() {
  var m = {'k1': 1, 'k2': 2};
  var s = m.entries.fold('', (acc, e) {
    if (acc.isEmpty) return '${e.key}:${e.value}';
    return '$acc,${e.key}:${e.value}';
  });
  print(s.contains('k1:1'));
  print(s.contains('k2:2'));
}"#,
        ["true", "true"]
    };

    remove_where_by_value_parity => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3, 'd': 4};
  m.removeWhere((k, v) => v.isEven);
  print(m.length);
  print(m.keys.join(','));
}"#,
        ["2", "a,c"]
    };

    remove_where_by_key_prefix => {
        r#"void main() {
  var m = {'temp_a': 1, 'keep_b': 2, 'temp_c': 3};
  m.removeWhere((k, v) => k.startsWith('temp_'));
  print(m.length);
  print(m.containsKey('keep_b'));
}"#,
        ["1", "true"]
    };

    remove_where_by_combined_key_and_value => {
        r#"void main() {
  var m = {'x': 10, 'y': 20, 'z': 5};
  m.removeWhere((k, v) => k == 'y' || v < 10);
  print(m.length);
  print(m.keys.join(','));
}"#,
        ["1", "y"]
    };

    remove_where_on_empty_map_is_noop => {
        r#"void main() {
  var m = <String, int>{};
  m.removeWhere((k, v) => true);
  print(m.isEmpty);
  print(m.length);
}"#,
        ["true", "0"]
    };

    remove_where_keeps_all_when_predicate_false => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  m.removeWhere((k, v) => false);
  print(m.length);
}"#,
        ["2"]
    };

    remove_where_then_entries_fold_recomputes_sum => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3, 'd': 4};
  m.removeWhere((k, v) => v > 2);
  var sum = m.entries.fold(0, (s, e) => s + e.value);
  print(sum);
  print(m.length);
}"#,
        ["3", "2"]
    };

    map_update_increments_existing_value => {
        r#"void main() {
  var m = {'count': 5};
  m.update('count', (v) => v + 1);
  print(m['count']);
}"#,
        ["6"]
    };

    map_update_with_if_absent_inserts_default => {
        r#"void main() {
  var m = <String, int>{};
  m.update('new', (v) => v + 10, ifAbsent: () => 0);
  print(m['new']);
  print(m.length);
}"#,
        ["0", "1"]
    };

    map_update_if_absent_not_called_when_key_exists => {
        r#"void main() {
  var m = {'x': 7};
  m.update('x', (v) => v * 2, ifAbsent: () => 100);
  print(m['x']);
}"#,
        ["14"]
    };

    map_update_all_doubles_every_value => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  m.updateAll((k, v) => v * 2);
  print(m.values.join(','));
}"#,
        ["2,4,6"]
    };

    map_update_all_uses_key_in_computation => {
        r#"void main() {
  var m = {'a': 0, 'b': 0};
  m.updateAll((k, v) => k.length);
  print(m['a']);
  print(m['b']);
}"#,
        ["1", "1"]
    };

    put_if_absent_chain_builds_nested_counter => {
        r#"void main() {
  var m = <String, Map<String, int>>{};
  m.putIfAbsent('users', () => {})..putIfAbsent('alice', () => 0);
  m['users']!['alice'] = m['users']!['alice']! + 1;
  print(m['users']!['alice']);
}"#,
        ["1"]
    };

    put_if_absent_chain_on_same_key_returns_existing => {
        r#"void main() {
  var m = <String, int>{};
  var a = m.putIfAbsent('k', () => 10);
  var b = m.putIfAbsent('k', () => 99);
  print(a);
  print(b);
  print(m['k']);
}"#,
        ["10", "10", "10"]
    };

    put_if_absent_sequential_keys_builds_map => {
        r#"void main() {
  var m = <String, int>{};
  m.putIfAbsent('a', () => 1);
  m.putIfAbsent('b', () => 2);
  m.putIfAbsent('c', () => 3);
  print(m.length);
  print(m.values.fold(0, (s, v) => s + v));
}"#,
        ["3", "6"]
    };

    put_if_absent_lazy_factory_not_called_when_present => {
        r#"void main() {
  var m = {'exists': 42};
  var called = 0;
  m.putIfAbsent('exists', () { called++; return 0; });
  print(m['exists']);
  print(called);
}"#,
        ["42", "0"]
    };

    entries_where_then_map_to_values => {
        r#"void main() {
  var m = {'a': 1, 'b': 10, 'c': 3, 'd': 12};
  var picked = m.entries
      .where((e) => e.value > 5)
      .map((e) => e.value)
      .toList()
    ..sort();
  print(picked.join(','));
}"#,
        ["10,12"]
    };

    entries_map_to_entry_swaps_key_value => {
        r#"void main() {
  var m = {'one': 1, 'two': 2};
  var swapped = Map.fromEntries(
    m.entries.map((e) => MapEntry('${e.value}', e.key.length)),
  );
  print(swapped['1']);
  print(swapped['2']);
}"#,
        ["3", "3"]
    };

    map_map_transforms_to_new_typed_map => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  var next = m.map((k, v) => MapEntry(k.toUpperCase(), '$v'));
  print(next['A']);
  print(next['B']);
}"#,
        ["1", "2"]
    };

    remove_where_then_put_if_absent_restores => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  m.removeWhere((k, v) => k == 'b');
  m.putIfAbsent('d', () => 4);
  print(m.length);
  print(m.containsKey('b'));
  print(m['d']);
}"#,
        ["3", "false", "4"]
    };

    update_chain_on_same_key_accumulates => {
        r#"void main() {
  var m = {'score': 0};
  m.update('score', (v) => v + 5);
  m.update('score', (v) => v + 3);
  m.update('score', (v) => v + 2);
  print(m['score']);
}"#,
        ["10"]
    };

    entries_fold_groups_by_value_bucket => {
        r#"void main() {
  var m = {'a': 1, 'b': 1, 'c': 2};
  var ones = m.entries.fold(0, (c, e) => e.value == 1 ? c + 1 : c);
  print(ones);
}"#,
        ["2"]
    };

    remove_where_value_less_than_keeps_high => {
        r#"void main() {
  var m = {'w': 100, 'x': 5, 'y': 50, 'z': 1};
  m.removeWhere((k, v) => v < 10);
  print(m.length);
  print(m.keys.join(','));
}"#,
        ["2", "w,y"]
    };

    entries_map_for_map_comprehension_style => {
        r#"void main() {
  var m = {'a': 2, 'b': 3};
  var cubed = {for (var e in m.entries) e.key: e.value * e.value * e.value};
  print(cubed['a']);
  print(cubed['b']);
}"#,
        ["8", "27"]
    };

    put_if_absent_with_update_combined_workflow => {
        r#"void main() {
  var m = <String, int>{};
  m.putIfAbsent('visits', () => 0);
  m.update('visits', (v) => v + 1);
  m.update('visits', (v) => v + 1);
  print(m['visits']);
}"#,
        ["2"]
    };

    remove_where_on_single_entry_map => {
        r#"void main() {
  var m = {'only': 42};
  m.removeWhere((k, v) => v == 42);
  print(m.isEmpty);
  print(m.length);
}"#,
        ["true", "0"]
    };

    entries_fold_finds_longest_key => {
        r#"void main() {
  var m = {'a': 1, 'bb': 2, 'ccc': 3};
  var longest = m.entries.fold('', (best, e) => e.key.length > best.length ? e.key : best);
  print(longest);
}"#,
        ["ccc"]
    };

    map_update_all_then_remove_where_pipeline => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  m.updateAll((k, v) => v * 10);
  m.removeWhere((k, v) => v == 20);
  print(m.length);
  print(m['a']);
  print(m['c']);
}"#,
        ["2", "10", "30"]
    };

    entries_map_to_bool_flags => {
        r#"void main() {
  var m = {'a': 1, 'b': 5, 'c': 3};
  var flags = m.entries.map((e) => e.value > 2).toList();
  print(flags.join(','));
}"#,
        ["false,true,true"]
    };

    put_if_absent_chain_for_default_lists => {
        r#"void main() {
  var m = <String, List<int>>{};
  m.putIfAbsent('nums', () => []).add(1);
  m.putIfAbsent('nums', () => []).add(2);
  print(m['nums']!.join(','));
  print(m['nums']!.length);
}"#,
        ["1,2", "2"]
    };

    remove_where_key_length_filter => {
        r#"void main() {
  var m = {'a': 1, 'bb': 2, 'ccc': 3};
  m.removeWhere((k, v) => k.length > 1);
  print(m.length);
  print(m['a']);
}"#,
        ["1", "1"]
    };

    entries_fold_product_of_values => {
        r#"void main() {
  var m = {'a': 2, 'b': 3, 'c': 4};
  var product = m.entries.fold(1, (p, e) => p * e.value);
  print(product);
}"#,
        ["24"]
    };

    update_with_if_absent_then_update_chain => {
        r#"void main() {
  var m = <String, int>{};
  m.update('token', (v) => v + 1, ifAbsent: () => 0);
  m.update('token', (v) => v + 5);
  print(m['token']);
}"#,
        ["5"]
    };

    entries_where_key_contains_substring => {
        r#"void main() {
  var m = {'user_alice': 1, 'admin_bob': 2, 'user_carol': 3};
  var userCount = m.entries.where((e) => e.key.contains('user_')).length;
  print(userCount);
}"#,
        ["2"]
    };

    map_from_entries_after_entries_map_filter => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3, 'd': 4};
  var evens = Map.fromEntries(
    m.entries.where((e) => e.value.isEven).map((e) => MapEntry(e.key, e.value * 10)),
  );
  print(evens.length);
  print(evens['b']);
  print(evens['d']);
}"#,
        ["2", "20", "40"]
    };

    remove_where_all_matching_leaves_empty => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  m.removeWhere((k, v) => v > 0);
  print(m.isEmpty);
}"#,
        ["true"]
    };

    put_if_absent_three_tier_grouping => {
        r#"void main() {
  var m = <String, Map<String, int>>{};
  m.putIfAbsent('g', () => {});
  m['g']!.putIfAbsent('x', () => 0);
  m['g']!.putIfAbsent('y', () => 0);
  m['g']!['x'] = 5;
  print(m['g']!['x']);
  print(m['g']!['y']);
}"#,
        ["5", "0"]
    };

    entries_fold_min_entry_by_value => {
        r#"void main() {
  var m = {'a': 30, 'b': 10, 'c': 20};
  var minKey = m.entries.fold('a', (best, e) => e.value < m[best]! ? e.key : best);
  print(minKey);
  print(m[minKey]);
}"#,
        ["b", "10"]
    };

    update_all_then_entries_map_snapshot => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  m.updateAll((k, v) => v + 100);
  var snapshot = m.entries.map((e) => e.value).toList()..sort();
  print(snapshot.join(','));
}"#,
        ["101,102"]
    };

    remove_where_preserves_insertion_order_of_remainder => {
        r#"void main() {
  var m = {'z': 1, 'a': 2, 'm': 3, 'b': 4};
  m.removeWhere((k, v) => v.isEven);
  print(m.keys.join(','));
}"#,
        ["z,m"]
    };

    entries_map_join_with_custom_separator => {
        r#"void main() {
  var m = {'k1': 1, 'k2': 2};
  var line = m.entries.map((e) => '${e.key}:${e.value}').join('; ');
  print(line);
}"#,
        ["k1:1; k2:2"]
    };
}
