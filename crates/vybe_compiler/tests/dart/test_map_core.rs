//! Core Map behaviors: literals, indexing, length, views, mutation, factories.

dart_cases! {
    map_literal_string_keys_and_int_values => {
        r#"void main() {
  var m = {'alpha': 1, 'beta': 2};
  print(m['alpha']);
  print(m['beta']);
}"#,
        ["1", "2"]
    };

    map_bracket_read_returns_null_for_missing_key => {
        r#"void main() {
  var m = {'a': 1};
  print(m['missing']);
}"#,
        ["null"]
    };

    map_bracket_assignment_inserts_new_entry => {
        r#"void main() {
  var m = <String, int>{};
  m['x'] = 10;
  print(m['x']);
  print(m.length);
}"#,
        ["10", "1"]
    };

    map_bracket_assignment_overwrites_existing_value => {
        r#"void main() {
  var m = {'k': 1};
  m['k'] = 99;
  print(m['k']);
  print(m.length);
}"#,
        ["99", "1"]
    };

    map_typed_empty_literal_has_zero_length => {
        r#"void main() {
  var m = <String, int>{};
  print(m.length);
  print(m.isEmpty);
}"#,
        ["0", "true"]
    };

    map_length_reflects_insertion_count => {
        r#"void main() {
  var m = <String, int>{};
  m['a'] = 1;
  m['b'] = 2;
  m['c'] = 3;
  print(m.length);
}"#,
        ["3"]
    };

    map_is_empty_true_on_unpopulated_map => {
        r#"void main() {
  var m = <int, String>{};
  print(m.isEmpty);
}"#,
        ["true"]
    };

    map_is_not_empty_after_first_insert => {
        r#"void main() {
  var m = <String, int>{};
  m['only'] = 1;
  print(m.isNotEmpty);
}"#,
        ["true"]
    };

    map_keys_to_list_preserves_insertion_order => {
        r#"void main() {
  var m = {'z': 1, 'a': 2, 'm': 3};
  var keys = m.keys.toList();
  print(keys.join(','));
}"#,
        ["z,a,m"]
    };

    map_values_to_list_matches_insertion_order => {
        r#"void main() {
  var m = {'a': 10, 'b': 20, 'c': 30};
  var vals = m.values.toList();
  print(vals.join(','));
}"#,
        ["10,20,30"]
    };

    map_entries_count_equals_length => {
        r#"void main() {
  var m = {'x': 1, 'y': 2, 'z': 3};
  print(m.entries.length);
}"#,
        ["3"]
    };

    map_add_inserts_new_key_and_returns_null => {
        r#"void main() {
  var m = <String, int>{};
  var prev = m.add('n', 7);
  print(prev);
  print(m['n']);
}"#,
        ["null", "7"]
    };

    map_add_on_existing_key_returns_previous_value => {
        r#"void main() {
  var m = {'k': 5};
  var prev = m.add('k', 9);
  print(prev);
  print(m['k']);
}"#,
        ["5", "9"]
    };

    map_add_all_merges_entries_from_other_map => {
        r#"void main() {
  var m = {'a': 1};
  m.addAll({'b': 2, 'c': 3});
  print(m.length);
  print(m['b']);
  print(m['c']);
}"#,
        ["3", "2", "3"]
    };

    map_add_all_overwrites_conflicting_keys => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  m.addAll({'b': 20, 'c': 3});
  print(m['b']);
  print(m.length);
}"#,
        ["20", "3"]
    };

    map_remove_existing_key_returns_removed_value => {
        r#"void main() {
  var m = {'a': 10, 'b': 20};
  var removed = m.remove('a');
  print(removed);
  print(m.containsKey('a'));
}"#,
        ["10", "false"]
    };

    map_remove_missing_key_returns_null => {
        r#"void main() {
  var m = {'a': 1};
  var removed = m.remove('ghost');
  print(removed);
}"#,
        ["null"]
    };

    map_remove_decrements_length => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  m.remove('a');
  print(m.length);
}"#,
        ["1"]
    };

    map_contains_key_true_when_present => {
        r#"void main() {
  var m = {'found': 42};
  print(m.containsKey('found'));
}"#,
        ["true"]
    };

    map_contains_key_false_when_absent => {
        r#"void main() {
  var m = {'a': 1};
  print(m.containsKey('z'));
}"#,
        ["false"]
    };

    map_contains_value_true_when_value_exists => {
        r#"void main() {
  var m = {'a': 100, 'b': 200};
  print(m.containsValue(200));
}"#,
        ["true"]
    };

    map_contains_value_false_when_value_missing => {
        r#"void main() {
  var m = {'a': 1};
  print(m.containsValue(999));
}"#,
        ["false"]
    };

    map_put_if_absent_inserts_when_key_missing => {
        r#"void main() {
  var m = <String, int>{};
  var v = m.putIfAbsent('new', () => 55);
  print(v);
  print(m['new']);
}"#,
        ["55", "55"]
    };

    map_put_if_absent_keeps_existing_without_recomputing => {
        r#"void main() {
  var m = {'keep': 11};
  var v = m.putIfAbsent('keep', () => 99);
  print(v);
  print(m['keep']);
}"#,
        ["11", "11"]
    };

    map_update_mutates_existing_value_in_place => {
        r#"void main() {
  var m = {'score': 10};
  m.update('score', (v) => v + 5);
  print(m['score']);
}"#,
        ["15"]
    };

    map_update_if_absent_inserts_when_key_missing => {
        r#"void main() {
  var m = <String, int>{};
  m.update('created', (v) => v + 1, ifAbsent: () => 0);
  print(m['created']);
}"#,
        ["0"]
    };

    map_update_all_applies_function_to_every_value => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  m.updateAll((k, v) => v * 10);
  print(m['a']);
  print(m['b']);
  print(m['c']);
}"#,
        ["10", "20", "30"]
    };

    map_foreach_visits_each_key_value_pair => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  m.forEach((k, v) => print('$k=$v'));
}"#,
        ["a=1", "b=2"]
    };

    map_map_transforms_entries_into_new_map => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  var doubled = m.map((k, v) => MapEntry(k, v * 2));
  print(doubled['a']);
  print(doubled['b']);
}"#,
        ["2", "4"]
    };

    map_clear_removes_all_entries => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  m.clear();
  print(m.length);
  print(m.isEmpty);
}"#,
        ["0", "true"]
    };

    map_from_copies_source_entries => {
        r#"void main() {
  var src = {'x': 1, 'y': 2};
  var copy = Map<String, int>.from(src);
  print(copy['x']);
  print(copy.length);
}"#,
        ["1", "2"]
    };

    map_from_entries_builds_map_from_entry_list => {
        r#"void main() {
  var m = Map.fromEntries([
    MapEntry('one', 1),
    MapEntry('two', 2),
  ]);
  print(m['one']);
  print(m['two']);
  print(m.length);
}"#,
        ["1", "2", "2"]
    };

    map_from_iterables_zips_parallel_key_value_lists => {
        r#"void main() {
  var m = Map.fromIterables(['a', 'b', 'c'], [10, 20, 30]);
  print(m['a']);
  print(m['c']);
  print(m.length);
}"#,
        ["10", "30", "3"]
    };

    map_of_factory_from_literal_entries => {
        r#"void main() {
  var m = Map<String, int>.of({'p': 3, 'q': 4});
  print(m['p']);
  print(m.length);
}"#,
        ["3", "2"]
    };

    map_identity_constructor_starts_empty => {
        r#"void main() {
  var m = Map.identity();
  m['obj'] = 1;
  print(m.length);
  print(m['obj']);
}"#,
        ["1", "1"]
    };

    map_cast_returns_typed_view_with_same_entries => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  var typed = m.cast<String, int>();
  print(typed['a']);
  print(typed.length);
}"#,
        ["1", "2"]
    };

    map_remove_where_deletes_matching_entries => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  m.removeWhere((k, v) => v.isEven);
  print(m.length);
  print(m.containsKey('b'));
  print(m.containsKey('a'));
}"#,
        ["2", "false", "true"]
    };

    map_entries_map_transforms_to_new_iterable => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  var labels = m.entries.map((e) => '${e.key}:${e.value}').toList();
  print(labels.join('|'));
}"#,
        ["a:1|b:2"]
    };

    map_stores_and_retrieves_null_values => {
        r#"void main() {
  var m = <String, int?>{'nullable': null};
  print(m.containsKey('nullable'));
  print(m['nullable']);
}"#,
        ["true", "null"]
    };

    map_int_keys_via_bracket_operator => {
        r#"void main() {
  var m = <int, String>{};
  m[1] = 'one';
  m[2] = 'two';
  print(m[1]);
  print(m[2]);
}"#,
        ["one", "two"]
    };

    map_nested_map_bracket_access => {
        r#"void main() {
  var m = {'outer': {'inner': 42}};
  print(m['outer']['inner']);
}"#,
        ["42"]
    };

    map_constructor_default_creates_empty_mutable_map => {
        r#"void main() {
  var m = Map<String, int>();
  m['k'] = 5;
  print(m.length);
  print(m['k']);
}"#,
        ["1", "5"]
    };

    map_literal_with_single_entry => {
        r#"void main() {
  var m = {'solo': 7};
  print(m.length);
  print(m['solo']);
}"#,
        ["1", "7"]
    };

    map_add_all_from_empty_map_is_noop => {
        r#"void main() {
  var m = {'a': 1};
  m.addAll(<String, int>{});
  print(m.length);
  print(m['a']);
}"#,
        ["1", "1"]
    };

    map_values_join_after_mutation => {
        r#"void main() {
  var m = <String, int>{};
  m['x'] = 3;
  m['y'] = 4;
  print(m.values.toList().reduce((a, b) => a + b));
}"#,
        ["7"]
    };

    map_keys_contains_after_bracket_write => {
        r#"void main() {
  var m = <String, int>{};
  m['fresh'] = 1;
  print(m.keys.contains('fresh'));
}"#,
        ["true"]
    };

    map_update_returns_final_value_from_updater => {
        r#"void main() {
  var m = {'n': 2};
  var result = m.update('n', (v) => v * 3);
  print(result);
  print(m['n']);
}"#,
        ["6", "6"]
    };

    map_foreach_accumulates_values => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  var sum = 0;
  m.forEach((k, v) => sum += v);
  print(sum);
}"#,
        ["6"]
    };

    map_from_iterables_with_int_keys => {
        r#"void main() {
  var m = Map<int, String>.fromIterables([0, 1, 2], ['zero', 'one', 'two']);
  print(m[1]);
  print(m.length);
}"#,
        ["one", "3"]
    };

    map_remove_where_on_empty_map_is_noop => {
        r#"void main() {
  var m = <String, int>{};
  m.removeWhere((k, v) => true);
  print(m.isEmpty);
  print(m.length);
}"#,
        ["true", "0"]
    };

    map_entries_first_key_value_after_literal => {
        r#"void main() {
  var m = {'first': 10, 'second': 20};
  var e = m.entries.first;
  print(e.key);
  print(e.value);
}"#,
        ["first", "10"]
    };

    map_add_all_then_remove_restores_partial_state => {
        r#"void main() {
  var m = {'a': 1};
  m.addAll({'b': 2});
  m.remove('a');
  print(m.length);
  print(m['b']);
  print(m.containsKey('a'));
}"#,
        ["1", "2", "false"]
    };
}
