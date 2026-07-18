//! `SplayTreeMap<K,V>` — comparison-ordered map backed by the shared sorted core.
//! Keys / values / entries enumerate in ascending key order; lookup is unchanged.

dart_cases! {
    splay_tree_map_keys_ascending => {
        r#"void main() {
  var m = SplayTreeMap<int, String>();
  m[3] = "c";
  m[1] = "a";
  m[2] = "b";
  for (var k in m.keys) {
    print(k);
  }
}"#,
        ["1", "2", "3"]
    };

    splay_tree_map_values_follow_key_order => {
        r#"void main() {
  var m = SplayTreeMap<int, String>();
  m[3] = "c";
  m[1] = "a";
  m[2] = "b";
  for (var v in m.values) {
    print(v);
  }
}"#,
        ["a", "b", "c"]
    };

    splay_tree_map_length_counts_entries => {
        r#"void main() {
  var m = SplayTreeMap<int, int>();
  m[10] = 1;
  m[20] = 2;
  m[30] = 3;
  print(m.length);
}"#,
        ["3"]
    };

    splay_tree_map_index_reads_value => {
        r#"void main() {
  var m = SplayTreeMap<int, String>();
  m[5] = "five";
  print(m[5]);
}"#,
        ["five"]
    };

    splay_tree_map_contains_key => {
        r#"void main() {
  var m = SplayTreeMap<int, String>();
  m[2] = "two";
  print(m.containsKey(2));
  print(m.containsKey(9));
}"#,
        ["true", "false"]
    };

    splay_tree_map_string_keys_sort_lexicographically => {
        r#"void main() {
  var m = SplayTreeMap<String, int>();
  m["zebra"] = 1;
  m["apple"] = 2;
  m["mango"] = 3;
  print(m.keys.toList());
}"#,
        [r#"[apple, mango, zebra]"#]
    };

    splay_tree_map_multidigit_keys_sort_numerically => {
        r#"void main() {
  var m = SplayTreeMap<int, int>();
  m[10] = 1;
  m[2] = 2;
  m[30] = 3;
  m[1] = 4;
  print(m.keys.toList());
}"#,
        ["[1, 2, 10, 30]"]
    };

    splay_tree_map_insert_out_of_order_still_sorts => {
        r#"void main() {
  var m = SplayTreeMap<int, int>();
  m[30] = 3;
  m[10] = 1;
  m[20] = 2;
  var sum = 0;
  for (var k in m.keys) {
    sum += k;
  }
  print(sum);
}"#,
        ["60"]
    };
}
