//! Unmodifiable List, Map, and Set: read access works; mutations are blocked.

dart_cases! {
    unmod_list_length_reports_source_size => {
        r#"void main() {
  var frozen = List.unmodifiable([10, 20, 30]);
  print(frozen.length);
}"#,
        ["3"]
    };

    unmod_list_index_read_returns_element => {
        r#"void main() {
  var frozen = List.unmodifiable(['a', 'b', 'c']);
  print(frozen[1]);
}"#,
        ["b"]
    };

    unmod_list_first_and_last_readable => {
        r#"void main() {
  var frozen = List.unmodifiable([5, 6, 7]);
  print(frozen.first);
  print(frozen.last);
}"#,
        ["5", "7"]
    };

    unmod_list_join_preserves_order => {
        r#"void main() {
  var frozen = List.unmodifiable([1, 2, 3]);
  print(frozen.join('-'));
}"#,
        ["1-2-3"]
    };

    unmod_list_contains_checks_membership => {
        r#"void main() {
  var frozen = List.unmodifiable([4, 8, 12]);
  print(frozen.contains(8));
  print(frozen.contains(99));
}"#,
        ["true", "false"]
    };

    unmod_list_is_empty_on_empty_source => {
        r#"void main() {
  var frozen = List.unmodifiable(<int>[]);
  print(frozen.isEmpty);
  print(frozen.isNotEmpty);
}"#,
        ["true", "false"]
    };

    unmod_list_reversed_view_does_not_mutate_source => {
        r#"void main() {
  var frozen = List.unmodifiable([1, 2, 3]);
  print(frozen.reversed.join(','));
  print(frozen.join(','));
}"#,
        ["3,2,1", "1,2,3"]
    };

    unmod_list_map_transforms_without_mutation => {
        r#"void main() {
  var frozen = List.unmodifiable([1, 2, 3]);
  print(frozen.map((n) => n * 2).join(','));
  print(frozen.length);
}"#,
        ["2,4,6", "3"]
    };

    unmod_list_where_filters_without_mutation => {
        r#"void main() {
  var frozen = List.unmodifiable([1, 2, 3, 4]);
  print(frozen.where((n) => n.isEven).join(','));
}"#,
        ["2,4"]
    };

    unmod_list_fold_accumulates_values => {
        r#"void main() {
  var frozen = List.unmodifiable([1, 2, 3]);
  print(frozen.fold(0, (a, b) => a + b));
}"#,
        ["6"]
    };

    unmod_list_index_of_finds_element => {
        r#"void main() {
  var frozen = List.unmodifiable([10, 20, 30]);
  print(frozen.indexOf(20));
  print(frozen.lastIndexOf(20));
}"#,
        ["1", "1"]
    };

    unmod_list_from_growable_source_captures_snapshot => {
        r#"void main() {
  var src = [1, 2];
  var frozen = List.unmodifiable(src);
  src.add(3);
  print(frozen.length);
  print(frozen.join(','));
}"#,
        ["2", "1,2"]
    };

    unmod_list_add_attempt_is_blocked => {
        r#"void main() {
  var frozen = List.unmodifiable([1, 2]);
  try { frozen.add(3); } catch (_) { print('blocked'); }
  print(frozen.length);
}"#,
        ["blocked", "2"]
    };

    unmod_list_add_all_attempt_is_blocked => {
        r#"void main() {
  var frozen = List.unmodifiable([1]);
  try { frozen.addAll([2, 3]); } catch (_) { print('blocked'); }
  print(frozen.join(','));
}"#,
        ["blocked", "1"]
    };

    unmod_list_insert_attempt_is_blocked => {
        r#"void main() {
  var frozen = List.unmodifiable([1, 3]);
  try { frozen.insert(1, 2); } catch (_) { print('blocked'); }
  print(frozen.join(','));
}"#,
        ["blocked", "1,3"]
    };

    unmod_list_index_assignment_is_blocked => {
        r#"void main() {
  var frozen = List.unmodifiable([1, 2, 3]);
  try { frozen[1] = 99; } catch (_) { print('blocked'); }
  print(frozen[1]);
}"#,
        ["blocked", "2"]
    };

    unmod_list_remove_by_value_is_blocked => {
        r#"void main() {
  var frozen = List.unmodifiable([1, 2, 3]);
  try { frozen.remove(2); } catch (_) { print('blocked'); }
  print(frozen.length);
}"#,
        ["blocked", "3"]
    };

    unmod_list_remove_at_is_blocked => {
        r#"void main() {
  var frozen = List.unmodifiable([10, 20]);
  try { frozen.removeAt(0); } catch (_) { print('blocked'); }
  print(frozen.first);
}"#,
        ["blocked", "10"]
    };

    unmod_list_remove_last_is_blocked => {
        r#"void main() {
  var frozen = List.unmodifiable([1, 2]);
  try { frozen.removeLast(); } catch (_) { print('blocked'); }
  print(frozen.last);
}"#,
        ["blocked", "2"]
    };

    unmod_list_remove_where_is_blocked => {
        r#"void main() {
  var frozen = List.unmodifiable([1, 2, 3, 4]);
  try { frozen.removeWhere((n) => n.isEven); } catch (_) { print('blocked'); }
  print(frozen.length);
}"#,
        ["blocked", "4"]
    };

    unmod_list_clear_is_blocked => {
        r#"void main() {
  var frozen = List.unmodifiable([1, 2]);
  try { frozen.clear(); } catch (_) { print('blocked'); }
  print(frozen.isEmpty);
}"#,
        ["blocked", "false"]
    };

    unmod_list_sort_is_blocked => {
        r#"void main() {
  var frozen = List.unmodifiable([3, 1, 2]);
  try { frozen.sort(); } catch (_) { print('blocked'); }
  print(frozen.join(','));
}"#,
        ["blocked", "3,1,2"]
    };

    unmod_list_set_range_is_blocked => {
        r#"void main() {
  var frozen = List.unmodifiable([1, 2, 3]);
  try { frozen.setRange(0, 2, [9, 9]); } catch (_) { print('blocked'); }
  print(frozen.join(','));
}"#,
        ["blocked", "1,2,3"]
    };

    unmod_map_length_reports_entry_count => {
        r#"void main() {
  var frozen = Map.unmodifiable({'a': 1, 'b': 2});
  print(frozen.length);
}"#,
        ["2"]
    };

    unmod_map_bracket_read_returns_value => {
        r#"void main() {
  var frozen = Map.unmodifiable({'x': 10, 'y': 20});
  print(frozen['x']);
  print(frozen['y']);
}"#,
        ["10", "20"]
    };

    unmod_map_contains_key_and_value => {
        r#"void main() {
  var frozen = Map.unmodifiable({'k': 5});
  print(frozen.containsKey('k'));
  print(frozen.containsValue(5));
  print(frozen.containsKey('missing'));
}"#,
        ["true", "true", "false"]
    };

    unmod_map_keys_join_preserves_order => {
        r#"void main() {
  var frozen = Map.unmodifiable({'z': 1, 'a': 2, 'm': 3});
  print(frozen.keys.join(','));
}"#,
        ["z,a,m"]
    };

    unmod_map_values_join_matches_insertion => {
        r#"void main() {
  var frozen = Map.unmodifiable({'a': 10, 'b': 20, 'c': 30});
  print(frozen.values.join(','));
}"#,
        ["10,20,30"]
    };

    unmod_map_entries_map_to_labels => {
        r#"void main() {
  var frozen = Map.unmodifiable({'a': 1, 'b': 2});
  var labels = frozen.entries.map((e) => '${e.key}:${e.value}').toList();
  print(labels.join('|'));
}"#,
        ["a:1|b:2"]
    };

    unmod_map_foreach_reads_pairs => {
        r#"void main() {
  var frozen = Map.unmodifiable({'p': 1, 'q': 2});
  frozen.forEach((k, v) => print('$k=$v'));
}"#,
        ["p=1", "q=2"]
    };

    unmod_map_is_empty_on_empty_source => {
        r#"void main() {
  var frozen = Map.unmodifiable(<String, int>{});
  print(frozen.isEmpty);
}"#,
        ["true"]
    };

    unmod_map_from_mutable_source_captures_snapshot => {
        r#"void main() {
  var src = <String, int>{'a': 1};
  var frozen = Map.unmodifiable(src);
  src['b'] = 2;
  print(frozen.length);
  print(frozen.containsKey('b'));
}"#,
        ["1", "false"]
    };

    unmod_map_bracket_assignment_is_blocked => {
        r#"void main() {
  var frozen = Map.unmodifiable({'a': 1});
  try { frozen['a'] = 99; } catch (_) { print('blocked'); }
  print(frozen['a']);
}"#,
        ["blocked", "1"]
    };

    unmod_map_insert_new_key_is_blocked => {
        r#"void main() {
  var frozen = Map.unmodifiable({'a': 1});
  try { frozen['new'] = 2; } catch (_) { print('blocked'); }
  print(frozen.length);
}"#,
        ["blocked", "1"]
    };

    unmod_map_remove_is_blocked => {
        r#"void main() {
  var frozen = Map.unmodifiable({'a': 1, 'b': 2});
  try { frozen.remove('a'); } catch (_) { print('blocked'); }
  print(frozen.length);
}"#,
        ["blocked", "2"]
    };

    unmod_map_clear_is_blocked => {
        r#"void main() {
  var frozen = Map.unmodifiable({'a': 1});
  try { frozen.clear(); } catch (_) { print('blocked'); }
  print(frozen.isEmpty);
}"#,
        ["blocked", "false"]
    };

    unmod_map_add_is_blocked => {
        r#"void main() {
  var frozen = Map.unmodifiable(<String, int>{});
  try { frozen.add('k', 1); } catch (_) { print('blocked'); }
  print(frozen.length);
}"#,
        ["blocked", "0"]
    };

    unmod_map_add_all_is_blocked => {
        r#"void main() {
  var frozen = Map.unmodifiable({'a': 1});
  try { frozen.addAll({'b': 2}); } catch (_) { print('blocked'); }
  print(frozen.containsKey('b'));
}"#,
        ["blocked", "false"]
    };

    unmod_map_put_if_absent_is_blocked => {
        r#"void main() {
  var frozen = Map.unmodifiable({'a': 1});
  try { frozen.putIfAbsent('b', () => 2); } catch (_) { print('blocked'); }
  print(frozen.length);
}"#,
        ["blocked", "1"]
    };

    unmod_map_update_is_blocked => {
        r#"void main() {
  var frozen = Map.unmodifiable({'score': 10});
  try { frozen.update('score', (v) => v + 1); } catch (_) { print('blocked'); }
  print(frozen['score']);
}"#,
        ["blocked", "10"]
    };

    unmod_map_remove_where_is_blocked => {
        r#"void main() {
  var frozen = Map.unmodifiable({'a': 1, 'b': 2});
  try { frozen.removeWhere((k, v) => v.isEven); } catch (_) { print('blocked'); }
  print(frozen.length);
}"#,
        ["blocked", "2"]
    };

    unmod_map_update_all_is_blocked => {
        r#"void main() {
  var frozen = Map.unmodifiable({'a': 1});
  try { frozen.updateAll((k, v) => v + 1); } catch (_) { print('blocked'); }
  print(frozen['a']);
}"#,
        ["blocked", "1"]
    };

    unmod_set_length_reports_size => {
        r#"void main() {
  var frozen = Set.unmodifiable({1, 2, 3});
  print(frozen.length);
}"#,
        ["3"]
    };

    unmod_set_contains_checks_membership => {
        r#"void main() {
  var frozen = Set.unmodifiable({'x', 'y'});
  print(frozen.contains('x'));
  print(frozen.contains('z'));
}"#,
        ["true", "false"]
    };

    unmod_set_to_list_sorted_for_display => {
        r#"void main() {
  var frozen = Set.unmodifiable({3, 1, 2});
  var list = frozen.toList()..sort();
  print(list.join(','));
}"#,
        ["1,2,3"]
    };

    unmod_set_is_empty_on_empty_source => {
        r#"void main() {
  var frozen = Set.unmodifiable(<int>{});
  print(frozen.isEmpty);
}"#,
        ["true"]
    };

    unmod_set_lookup_deduplicates_source => {
        r#"void main() {
  var frozen = Set.unmodifiable({1, 2, 2, 3});
  print(frozen.length);
}"#,
        ["3"]
    };

    unmod_set_from_mutable_source_captures_snapshot => {
        r#"void main() {
  var src = {1, 2};
  var frozen = Set.unmodifiable(src);
  src.add(3);
  print(frozen.length);
}"#,
        ["2"]
    };

    unmod_set_add_is_blocked => {
        r#"void main() {
  var frozen = Set.unmodifiable({1, 2});
  try { frozen.add(3); } catch (_) { print('blocked'); }
  print(frozen.length);
}"#,
        ["blocked", "2"]
    };

    unmod_set_add_all_is_blocked => {
        r#"void main() {
  var frozen = Set.unmodifiable({1});
  try { frozen.addAll({2, 3}); } catch (_) { print('blocked'); }
  print(frozen.length);
}"#,
        ["blocked", "1"]
    };

    unmod_set_remove_is_blocked => {
        r#"void main() {
  var frozen = Set.unmodifiable({1, 2, 3});
  try { frozen.remove(2); } catch (_) { print('blocked'); }
  print(frozen.contains(2));
}"#,
        ["blocked", "true"]
    };

    unmod_set_remove_where_is_blocked => {
        r#"void main() {
  var frozen = Set.unmodifiable({1, 2, 3, 4});
  try { frozen.removeWhere((n) => n.isEven); } catch (_) { print('blocked'); }
  print(frozen.length);
}"#,
        ["blocked", "4"]
    };

    unmod_set_clear_is_blocked => {
        r#"void main() {
  var frozen = Set.unmodifiable({1, 2});
  try { frozen.clear(); } catch (_) { print('blocked'); }
  print(frozen.isEmpty);
}"#,
        ["blocked", "false"]
    };

    unmod_set_retain_where_is_blocked => {
        r#"void main() {
  var frozen = Set.unmodifiable({1, 2, 3});
  try { frozen.retainWhere((n) => n == 1); } catch (_) { print('blocked'); }
  print(frozen.length);
}"#,
        ["blocked", "3"]
    };
}
