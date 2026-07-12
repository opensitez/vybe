//! Linked hash insertion order: default maps, LinkedHashMap, mutations.

dart_cases! {
    linked_literal_map_keys_follow_insertion_order => {
        r#"void main() {
  var m = {'c': 3, 'a': 1, 'b': 2};
  print(m.keys.join(','));
}"#,
        ["c,a,b"]
    };

    linked_literal_map_values_follow_insertion_order => {
        r#"void main() {
  var m = {'first': 10, 'second': 20, 'third': 30};
  print(m.values.join(','));
}"#,
        ["10,20,30"]
    };

    linked_empty_map_keys_join_is_empty => {
        r#"void main() {
  var m = <String, int>{};
  print(m.keys.join(','));
  print(m.keys.length);
}"#,
        ["", "0"]
    };

    linked_single_entry_key_order => {
        r#"void main() {
  var m = {'only': 1};
  print(m.keys.join(','));
  print(m.values.join(','));
}"#,
        ["only", "1"]
    };

    linked_explicit_linked_hash_map_constructor => {
        r#"void main() {
  var m = LinkedHashMap<String, int>();
  m['z'] = 1;
  m['a'] = 2;
  m['m'] = 3;
  print(m.keys.join(','));
}"#,
        ["z,a,m"]
    };

    linked_hash_map_from_empty_then_sequential_puts => {
        r#"void main() {
  var m = LinkedHashMap<int, String>();
  m[3] = 'three';
  m[1] = 'one';
  m[2] = 'two';
  print(m.keys.join(','));
  print(m.values.join(','));
}"#,
        ["3,1,2", "three,one,two"]
    };

    linked_hash_map_typed_constructor_preserves_order => {
        r#"void main() {
  var m = LinkedHashMap<String, String>();
  m['red'] = 'r';
  m['green'] = 'g';
  m['blue'] = 'b';
  print(m.keys.join('-'));
}"#,
        ["red-green-blue"]
    };

    linked_remove_middle_key_then_keys_join => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  m.remove('b');
  print(m.keys.join(','));
}"#,
        ["a,c"]
    };

    linked_remove_first_key_preserves_tail_order => {
        r#"void main() {
  var m = {'x': 1, 'y': 2, 'z': 3};
  m.remove('x');
  print(m.keys.join(','));
}"#,
        ["y,z"]
    };

    linked_remove_last_key_preserves_prefix_order => {
        r#"void main() {
  var m = {'x': 1, 'y': 2, 'z': 3};
  m.remove('z');
  print(m.keys.join(','));
}"#,
        ["x,y"]
    };

    linked_reinsert_removed_key_moves_to_end => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  m.remove('b');
  m['b'] = 20;
  print(m.keys.join(','));
}"#,
        ["a,c,b"]
    };

    linked_reinsert_first_key_after_remove => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  m.remove('a');
  m['a'] = 10;
  print(m.keys.join(','));
}"#,
        ["b,c,a"]
    };

    linked_update_existing_value_does_not_reorder_key => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  m['b'] = 99;
  print(m.keys.join(','));
  print(m['b']);
}"#,
        ["a,b,c", "99"]
    };

    linked_bracket_overwrite_keeps_key_position => {
        r#"void main() {
  var m = {'p': 1, 'q': 2};
  m['p'] = 100;
  print(m.keys.join(','));
  print(m.values.join(','));
}"#,
        ["p,q", "100,2"]
    };

    linked_add_all_appends_new_keys_in_argument_order => {
        r#"void main() {
  var m = {'a': 1};
  m.addAll({'x': 10, 'y': 20});
  print(m.keys.join(','));
}"#,
        ["a,x,y"]
    };

    linked_add_all_overwrites_without_reordering_existing => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  m.addAll({'a': 10, 'c': 3});
  print(m.keys.join(','));
  print(m['a']);
}"#,
        ["a,b,c", "10"]
    };

    linked_remove_where_preserves_relative_order_of_survivors => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3, 'd': 4};
  m.removeWhere((k, v) => v.isEven);
  print(m.keys.join(','));
}"#,
        ["a,c"]
    };

    linked_remove_where_on_all_but_one => {
        r#"void main() {
  var m = {'keep': 1, 'drop1': 2, 'drop2': 4};
  m.removeWhere((k, v) => v != 1);
  print(m.keys.join(','));
  print(m.length);
}"#,
        ["keep", "1"]
    };

    linked_entries_iteration_matches_key_order => {
        r#"void main() {
  var m = {'w': 1, 'x': 2, 'y': 3};
  var order = <String>[];
  for (var e in m.entries) { order.add(e.key); }
  print(order.join(','));
}"#,
        ["w,x,y"]
    };

    linked_entries_map_preserves_source_key_order => {
        r#"void main() {
  var m = {'c': 3, 'b': 2, 'a': 1};
  var keys = m.entries.map((e) => e.key).toList();
  print(keys.join(','));
}"#,
        ["c,b,a"]
    };

    linked_for_each_visits_in_insertion_order => {
        r#"void main() {
  var m = {'one': 1, 'two': 2, 'three': 3};
  var buf = <String>[];
  m.forEach((k, v) => buf.add(k));
  print(buf.join('|'));
}"#,
        ["one|two|three"]
    };

    linked_put_if_absent_appends_new_key_at_end => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  m.putIfAbsent('c', () => 3);
  print(m.keys.join(','));
}"#,
        ["a,b,c"]
    };

    linked_put_if_absent_on_existing_does_not_reorder => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  m.putIfAbsent('a', () => 99);
  print(m.keys.join(','));
  print(m['a']);
}"#,
        ["a,b", "1"]
    };

    linked_update_if_absent_inserts_at_end => {
        r#"void main() {
  var m = {'x': 1};
  m.update('y', (v) => v + 1, ifAbsent: () => 0);
  print(m.keys.join(','));
  print(m['y']);
}"#,
        ["x,y", "0"]
    };

    linked_clear_then_rebuild_new_order => {
        r#"void main() {
  var m = {'old1': 1, 'old2': 2};
  m.clear();
  m['new1'] = 10;
  m['new2'] = 20;
  print(m.keys.join(','));
}"#,
        ["new1,new2"]
    };

    linked_map_from_entries_preserves_list_order => {
        r#"void main() {
  var m = Map.fromEntries([
    MapEntry('third', 3),
    MapEntry('first', 1),
    MapEntry('second', 2),
  ]);
  print(m.keys.join(','));
}"#,
        ["third,first,second"]
    };

    linked_map_from_iterables_zips_in_parallel_order => {
        r#"void main() {
  var m = Map.fromIterables(['z', 'a', 'm'], [1, 2, 3]);
  print(m.keys.join(','));
  print(m.values.join(','));
}"#,
        ["z,a,m", "1,2,3"]
    };

    linked_hash_map_of_copies_source_order => {
        r#"void main() {
  var src = {'b': 2, 'a': 1, 'c': 3};
  var m = LinkedHashMap<String, int>.of(src);
  print(m.keys.join(','));
}"#,
        ["b,a,c"]
    };

    linked_hash_map_from_copies_source_order => {
        r#"void main() {
  var src = {'x': 10, 'y': 20};
  var m = LinkedHashMap<String, int>.from(src);
  print(m.keys.join(','));
}"#,
        ["x,y"]
    };

    linked_remove_and_add_all_rebuilds_tail => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  m.remove('b');
  m.addAll({'d': 4, 'e': 5});
  print(m.keys.join(','));
}"#,
        ["a,c,d,e"]
    };

    linked_multiple_reinserts_shuffle_end_positions => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  m.remove('a');
  m.remove('c');
  m['a'] = 10;
  m['c'] = 30;
  print(m.keys.join(','));
}"#,
        ["b,a,c"]
    };

    linked_numeric_string_keys_maintain_insertion => {
        r#"void main() {
  var m = <String, int>{};
  m['10'] = 10;
  m['2'] = 2;
  m['1'] = 1;
  print(m.keys.join(','));
}"#,
        ["10,2,1"]
    };

    linked_int_key_map_iteration_order => {
        r#"void main() {
  var m = <int, String>{};
  m[300] = 'c';
  m[100] = 'a';
  m[200] = 'b';
  print(m.keys.join(','));
}"#,
        ["300,100,200"]
    };

    linked_values_to_list_matches_key_order => {
        r#"void main() {
  var m = {'d': 4, 'b': 2, 'c': 3, 'a': 1};
  print(m.values.toList().join(','));
}"#,
        ["4,2,3,1"]
    };

    linked_keys_to_list_is_independent_copy => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  var keys = m.keys.toList();
  m['c'] = 3;
  print(keys.join(','));
  print(m.keys.join(','));
}"#,
        ["a,b", "a,b,c"]
    };

    linked_update_all_preserves_key_order => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  m.updateAll((k, v) => v * 10);
  print(m.keys.join(','));
  print(m.values.join(','));
}"#,
        ["a,b,c", "10,20,30"]
    };

    linked_map_add_returns_null_on_new_key_at_end => {
        r#"void main() {
  var m = {'a': 1};
  m.add('b', 2);
  print(m.keys.join(','));
}"#,
        ["a,b"]
    };

    linked_map_add_on_existing_preserves_order => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  m.add('a', 10);
  print(m.keys.join(','));
}"#,
        ["a,b"]
    };

    linked_literal_with_duplicate_keys_keeps_last_value => {
        r#"void main() {
  var m = {'k': 1, 'k': 2};
  print(m.keys.join(','));
  print(m['k']);
  print(m.length);
}"#,
        ["k", "2", "1"]
    };

    linked_spread_map_preserves_left_to_right_key_order => {
        r#"void main() {
  var a = {'x': 1, 'y': 2};
  var b = {'z': 3};
  var m = {...a, ...b};
  print(m.keys.join(','));
}"#,
        ["x,y,z"]
    };

    linked_spread_overwrite_keeps_first_key_position => {
        r#"void main() {
  var a = {'k': 1, 'm': 2};
  var b = {'k': 99, 'n': 3};
  var m = {...a, ...b};
  print(m.keys.join(','));
  print(m['k']);
}"#,
        ["k,m,n", "99"]
    };

    linked_remove_where_then_put_if_absent => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  m.removeWhere((k, v) => k == 'b');
  m.putIfAbsent('d', () => 4);
  print(m.keys.join(','));
}"#,
        ["a,c,d"]
    };

    linked_long_chain_of_appends => {
        r#"void main() {
  var m = <String, int>{};
  m['e'] = 5;
  m['d'] = 4;
  m['c'] = 3;
  m['b'] = 2;
  m['a'] = 1;
  print(m.keys.join(''));
}"#,
        ["edcba"]
    };

    linked_entries_fold_builds_ordered_key_string => {
        r#"void main() {
  var m = {'p': 1, 'q': 2, 'r': 3};
  var s = m.entries.fold('', (acc, e) => acc + e.key);
  print(s);
}"#,
        ["pqr"]
    };

    linked_hash_set_add_order_when_converted_to_list => {
        r#"void main() {
  var s = LinkedHashSet<String>();
  s.add('c');
  s.add('a');
  s.add('b');
  print(s.join(','));
}"#,
        ["c,a,b"]
    };

    linked_hash_set_remove_and_readd_moves_to_end => {
        r#"void main() {
  var s = LinkedHashSet<int>();
  s.add(1);
  s.add(2);
  s.add(3);
  s.remove(2);
  s.add(2);
  print(s.join(','));
}"#,
        ["1,3,2"]
    };
}
