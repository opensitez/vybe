//! `SplayTreeSet<T>` — comparison-ordered set backed by the shared sorted core.
//! Ascending iteration, `first`=min / `last`=max, and dedupe semantics.

dart_cases! {
    splay_tree_set_iterates_ascending => {
        r#"void main() {
  var s = SplayTreeSet<int>();
  s.add(5);
  s.add(1);
  s.add(3);
  for (var x in s) {
    print(x);
  }
}"#,
        ["1", "3", "5"]
    };

    splay_tree_set_insert_out_of_order_sorts => {
        r#"void main() {
  var s = SplayTreeSet<int>();
  s.add(30);
  s.add(10);
  s.add(20);
  print(s.toList());
}"#,
        ["[10, 20, 30]"]
    };

    splay_tree_set_first_is_minimum => {
        r#"void main() {
  var s = SplayTreeSet<int>();
  s.add(8);
  s.add(2);
  s.add(5);
  print(s.first);
}"#,
        ["2"]
    };

    splay_tree_set_last_is_maximum => {
        r#"void main() {
  var s = SplayTreeSet<int>();
  s.add(8);
  s.add(2);
  s.add(5);
  print(s.last);
}"#,
        ["8"]
    };

    splay_tree_set_dedupes_duplicates => {
        r#"void main() {
  var s = SplayTreeSet<int>();
  s.add(1);
  s.add(2);
  s.add(2);
  s.add(3);
  s.add(1);
  print(s.length);
}"#,
        ["3"]
    };

    splay_tree_set_contains_member => {
        r#"void main() {
  var s = SplayTreeSet<int>();
  s.add(4);
  s.add(9);
  print(s.contains(4));
  print(s.contains(7));
}"#,
        ["true", "false"]
    };

    splay_tree_set_add_returns_bool => {
        r#"void main() {
  var s = SplayTreeSet<int>();
  print(s.add(1));
  print(s.add(1));
}"#,
        ["true", "false"]
    };

    splay_tree_set_string_keys_sort_lexicographically => {
        r#"void main() {
  var s = SplayTreeSet<String>();
  s.add("zebra");
  s.add("apple");
  s.add("mango");
  print(s.toList());
}"#,
        [r#"[apple, mango, zebra]"#]
    };
}
