//! Expando property bags on objects: attach hidden data, null-key rejection,
//! multiple keys per object, and typed Expandos. WeakReference omitted (exotic).

dart_cases! {
    expando_set_and_get_string_value => {
        r#"void main() {
  final bag = Expando<String>();
  var obj = Object();
  bag[obj] = 'hello';
  print(bag[obj]);
}"#,
        ["hello"]
    };

    expando_get_before_set_returns_null => {
        r#"void main() {
  final bag = Expando<String>();
  var obj = Object();
  print(bag[obj] == null);
}"#,
        ["true"]
    };

    expando_overwrite_existing_value => {
        r#"void main() {
  final bag = Expando<int>();
  var obj = Object();
  bag[obj] = 1;
  bag[obj] = 2;
  print(bag[obj]);
}"#,
        ["2"]
    };

    expando_store_null_value_explicitly => {
        r#"void main() {
  final bag = Expando<String?>();
  var obj = Object();
  bag[obj] = null;
  print(bag[obj] == null);
}"#,
        ["true"]
    };

    expando_two_objects_independent_values => {
        r#"void main() {
  final bag = Expando<int>();
  var a = Object();
  var b = Object();
  bag[a] = 10;
  bag[b] = 20;
  print(bag[a]);
  print(bag[b]);
}"#,
        ["10", "20"]
    };

    expando_same_object_multiple_expandos => {
        r#"void main() {
  final names = Expando<String>();
  final scores = Expando<int>();
  var obj = Object();
  names[obj] = 'alice';
  scores[obj] = 99;
  print(names[obj]);
  print(scores[obj]);
}"#,
        ["alice", "99"]
    };

    expando_int_values => {
        r#"void main() {
  final bag = Expando<int>();
  var obj = Object();
  bag[obj] = 42;
  print(bag[obj]);
}"#,
        ["42"]
    };

    expando_bool_values => {
        r#"void main() {
  final bag = Expando<bool>();
  var obj = Object();
  bag[obj] = true;
  print(bag[obj]);
}"#,
        ["true"]
    };

    expando_double_values => {
        r#"void main() {
  final bag = Expando<double>();
  var obj = Object();
  bag[obj] = 3.14;
  print(bag[obj]);
}"#,
        ["3.14"]
    };

    expando_list_values => {
        r#"void main() {
  final bag = Expando<List<int>>();
  var obj = Object();
  bag[obj] = [1, 2, 3];
  print(bag[obj]!.join(','));
}"#,
        ["1,2,3"]
    };

    expando_map_values => {
        r#"void main() {
  final bag = Expando<Map<String, int>>();
  var obj = Object();
  bag[obj] = {'a': 1};
  print(bag[obj]!['a']);
}"#,
        ["1"]
    };

    expando_on_custom_class_instance => {
        r#"class Widget {
  int id;
  Widget(this.id);
}
void main() {
  final bag = Expando<String>();
  var w = Widget(5);
  bag[w] = 'panel';
  print(bag[w]);
}"#,
        ["panel"]
    };

    expando_on_two_custom_instances => {
        r#"class Node {
  String label;
  Node(this.label);
}
void main() {
  final bag = Expando<int>();
  var n1 = Node('a');
  var n2 = Node('b');
  bag[n1] = 1;
  bag[n2] = 2;
  print(bag[n1]);
  print(bag[n2]);
}"#,
        ["1", "2"]
    };

    expando_on_subclass_instance => {
        r#"class Animal {}
class Dog extends Animal {}
void main() {
  final bag = Expando<String>();
  var d = Dog();
  bag[d] = 'woof';
  print(bag[d]);
}"#,
        ["woof"]
    };

    expando_function_object_as_key => {
        r#"void helper() {}
void main() {
  final bag = Expando<String>();
  bag[helper] = 'fn';
  print(bag[helper]);
}"#,
        ["fn"]
    };

    expando_list_object_as_key => {
        r#"void main() {
  final bag = Expando<String>();
  var list = [1, 2];
  bag[list] = 'list-key';
  print(bag[list]);
}"#,
        ["list-key"]
    };

    expando_string_object_as_key => {
        r#"void main() {
  final bag = Expando<int>();
  var s = 'key';
  bag[s] = 7;
  print(bag[s]);
}"#,
        ["7"]
    };

    expando_distinct_objects_same_field_values => {
        r#"class Pair {
  int a;
  int b;
  Pair(this.a, this.b);
}
void main() {
  final bag = Expando<String>();
  var p1 = Pair(1, 2);
  var p2 = Pair(1, 2);
  bag[p1] = 'first';
  bag[p2] = 'second';
  print(bag[p1]);
  print(bag[p2]);
}"#,
        ["first", "second"]
    };

    expando_three_objects_three_entries => {
        r#"void main() {
  final bag = Expando<int>();
  var o1 = Object();
  var o2 = Object();
  var o3 = Object();
  bag[o1] = 1;
  bag[o2] = 2;
  bag[o3] = 3;
  print(bag[o1] + bag[o2]! + bag[o3]!);
}"#,
        ["6"]
    };

    expando_isolated_between_instances => {
        r#"void main() {
  final bag1 = Expando<String>();
  final bag2 = Expando<String>();
  var obj = Object();
  bag1[obj] = 'one';
  bag2[obj] = 'two';
  print(bag1[obj]);
  print(bag2[obj]);
}"#,
        ["one", "two"]
    };

    expando_null_key_on_set_throws => {
        r#"void main() {
  final bag = Expando<String>();
  try {
    bag[null] = 'bad';
    print('no throw');
  } catch (e) {
    print('threw');
  }
}"#,
        ["threw"]
    };

    expando_null_key_on_get_throws => {
        r#"void main() {
  final bag = Expando<String>();
  try {
    print(bag[null]);
    print('no throw');
  } catch (e) {
    print('threw');
  }
}"#,
        ["threw"]
    };

    expando_replace_after_initial_set => {
        r#"void main() {
  final bag = Expando<String>();
  var obj = Object();
  bag[obj] = 'old';
  bag[obj] = 'new';
  print(bag[obj]);
}"#,
        ["new"]
    };

    expando_clear_by_setting_null_then_get => {
        r#"void main() {
  final bag = Expando<String?>();
  var obj = Object();
  bag[obj] = 'present';
  bag[obj] = null;
  print(bag[obj] == null);
}"#,
        ["true"]
    };

    expando_on_same_reference_twice => {
        r#"void main() {
  final bag = Expando<int>();
  var obj = Object();
  var alias = obj;
  bag[obj] = 5;
  print(bag[alias]);
}"#,
        ["5"]
    };

    expando_chained_get_after_multiple_sets => {
        r#"void main() {
  final bag = Expando<int>();
  var obj = Object();
  bag[obj] = 1;
  bag[obj] = bag[obj]! + 1;
  bag[obj] = bag[obj]! + 1;
  print(bag[obj]);
}"#,
        ["3"]
    };

    expando_negative_int_value => {
        r#"void main() {
  final bag = Expando<int>();
  var obj = Object();
  bag[obj] = -99;
  print(bag[obj]);
}"#,
        ["-99"]
    };

    expando_zero_value => {
        r#"void main() {
  final bag = Expando<int>();
  var obj = Object();
  bag[obj] = 0;
  print(bag[obj]);
}"#,
        ["0"]
    };

    expando_empty_string_value => {
        r#"void main() {
  final bag = Expando<String>();
  var obj = Object();
  bag[obj] = '';
  print(bag[obj]!.length);
}"#,
        ["0"]
    };

    expando_long_string_value => {
        r#"void main() {
  final bag = Expando<String>();
  var obj = Object();
  bag[obj] = 'abcdefghij';
  print(bag[obj]!.length);
}"#,
        ["10"]
    };

    expando_multiple_keys_same_expandos_different_objects => {
        r#"void main() {
  final bag = Expando<String>();
  var keys = List.generate(3, (_) => Object());
  bag[keys[0]] = 'a';
  bag[keys[1]] = 'b';
  bag[keys[2]] = 'c';
  print(bag[keys[0]]! + bag[keys[1]]! + bag[keys[2]]!);
}"#,
        ["abc"]
    };

    expando_object_value_roundtrip => {
        r#"class Holder {
  int v;
  Holder(this.v);
}
void main() {
  final bag = Expando<Holder>();
  var key = Object();
  var holder = Holder(8);
  bag[key] = holder;
  print(bag[key]!.v);
}"#,
        ["8"]
    };

    expando_read_unrelated_object_returns_null => {
        r#"void main() {
  final bag = Expando<String>();
  var stored = Object();
  var other = Object();
  bag[stored] = 'data';
  print(bag[other] == null);
}"#,
        ["true"]
    };

    expando_set_on_one_does_not_affect_other_expando => {
        r#"void main() {
  final a = Expando<int>();
  final b = Expando<int>();
  var obj = Object();
  a[obj] = 100;
  print(b[obj] == null);
}"#,
        ["true"]
    };

    expando_widget_tree_metadata_pattern => {
        r#"class Element {
  int depth;
  Element(this.depth);
}
void main() {
  final meta = Expando<String>();
  var root = Element(0);
  var child = Element(1);
  meta[root] = 'root';
  meta[child] = 'child';
  print(meta[root]);
  print(meta[child]);
}"#,
        ["root", "child"]
    };

    expando_cache_pattern_lookup => {
        r#"class Service {
  int id;
  Service(this.id);
}
void main() {
  final cache = Expando<String>();
  var svc = Service(42);
  cache[svc] = 'cached';
  print(cache[svc]);
}"#,
        ["cached"]
    };

    expando_bool_false_stored => {
        r#"void main() {
  final bag = Expando<bool>();
  var obj = Object();
  bag[obj] = false;
  print(bag[obj]);
}"#,
        ["false"]
    };

    expando_increment_numeric_metadata => {
        r#"void main() {
  final visits = Expando<int>();
  var page = Object();
  visits[page] = 1;
  visits[page] = visits[page]! + 1;
  print(visits[page]);
}"#,
        ["2"]
    };

    expando_toggle_bool_metadata => {
        r#"void main() {
  final flags = Expando<bool>();
  var item = Object();
  flags[item] = true;
  flags[item] = !flags[item]!;
  print(flags[item]);
}"#,
        ["false"]
    };

    expando_sequential_objects_unique_slots => {
        r#"void main() {
  final bag = Expando<int>();
  var results = <int>[];
  for (var i = 0; i < 3; i++) {
    var obj = Object();
    bag[obj] = i;
    results.add(bag[obj]!);
  }
  print(results.join(','));
}"#,
        ["0,1,2"]
    };
}
