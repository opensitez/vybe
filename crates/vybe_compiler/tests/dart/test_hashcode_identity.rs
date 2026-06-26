//! Object identity via identical(), and hashCode stability for same and different objects.

dart_cases! {
    identical_same_list_variable_is_true => {
        r#"void main() {
  var list = [1, 2, 3];
  print(identical(list, list));
}"#,
        ["true"]
    };

    identical_list_alias_same_reference => {
        r#"void main() {
  var original = [1, 2];
  var alias = original;
  print(identical(original, alias));
}"#,
        ["true"]
    };

    identical_two_fresh_list_literals_is_false => {
        r#"void main() {
  print(identical([1], [1]));
}"#,
        ["false"]
    };

    identical_two_fresh_map_literals_is_false => {
        r#"void main() {
  print(identical({'a': 1}, {'a': 1}));
}"#,
        ["false"]
    };

    identical_same_map_variable_is_true => {
        r#"void main() {
  var map = {'k': 1};
  print(identical(map, map));
}"#,
        ["true"]
    };

    identical_map_alias_same_reference => {
        r#"void main() {
  var left = {'x': 9};
  var right = left;
  print(identical(left, right));
}"#,
        ["true"]
    };

    identical_same_int_variable_is_true => {
        r#"void main() {
  var n = 42;
  print(identical(n, n));
}"#,
        ["true"]
    };

    identical_same_string_variable_is_true => {
        r#"void main() {
  var s = 'dart';
  print(identical(s, s));
}"#,
        ["true"]
    };

    identical_null_with_null_is_true => {
        r#"void main() {
  print(identical(null, null));
}"#,
        ["true"]
    };

    identical_null_with_object_is_false => {
        r#"void main() {
  var value = 1;
  print(identical(null, value));
}"#,
        ["false"]
    };

    identical_two_distinct_string_instances => {
        r#"void main() {
  var a = String.fromCharCodes([97, 98]);
  var b = String.fromCharCodes([97, 98]);
  print(identical(a, b));
  print(a == b);
}"#,
        ["false", "true"]
    };

    identical_enum_same_member_is_true => {
        r#"enum Mode { on, off }
void main() {
  var a = Mode.on;
  var b = Mode.on;
  print(identical(a, b));
}"#,
        ["true"]
    };

    identical_enum_different_members_is_false => {
        r#"enum Mode { on, off }
void main() {
  print(identical(Mode.on, Mode.off));
}"#,
        ["false"]
    };

    identical_custom_class_same_instance => {
        r#"class Token {
  int id;
  Token(this.id);
}
void main() {
  var t = Token(1);
  print(identical(t, t));
}"#,
        ["true"]
    };

    identical_custom_class_two_instances_is_false => {
        r#"class Token {
  int id;
  Token(this.id);
}
void main() {
  print(identical(Token(1), Token(1)));
}"#,
        ["false"]
    };

    identical_set_alias_same_reference => {
        r#"void main() {
  var set = {1, 2};
  var alias = set;
  print(identical(set, alias));
}"#,
        ["true"]
    };

    identical_two_fresh_sets_is_false => {
        r#"void main() {
  print(identical({1, 2}, {1, 2}));
}"#,
        ["false"]
    };

    identical_closure_same_reference => {
        r#"void main() {
  int Function(int) fn = (n) => n + 1;
  print(identical(fn, fn));
}"#,
        ["true"]
    };

    identical_two_distinct_closures_is_false => {
        r#"void main() {
  print(identical((n) => n + 1, (n) => n + 1));
}"#,
        ["false"]
    };

    identical_after_reassign_keeps_reference => {
        r#"void main() {
  var box = [0];
  var holder = box;
  box = holder;
  print(identical(box, holder));
}"#,
        ["true"]
    };

    identical_bool_literals_same_value => {
        r#"void main() {
  print(identical(true, true));
  print(identical(false, false));
}"#,
        ["true", "true"]
    };

    identical_record_same_variable => {
        r#"void main() {
  var pair = (1, 'a');
  print(identical(pair, pair));
}"#,
        ["true"]
    };

    identical_two_record_literals_is_false => {
        r#"void main() {
  print(identical((1, 'a'), (1, 'a')));
}"#,
        ["false"]
    };

    identical_list_sublist_view_same_reference => {
        r#"void main() {
  var list = [1, 2, 3, 4];
  var view = list;
  print(identical(list, view));
}"#,
        ["true"]
    };

    hashcode_int_stable_on_same_value => {
        r#"void main() {
  var n = 99;
  print(n.hashCode == n.hashCode);
}"#,
        ["true"]
    };

    hashcode_string_stable_on_same_object => {
        r#"void main() {
  var s = 'stable';
  print(s.hashCode == s.hashCode);
}"#,
        ["true"]
    };

    hashcode_bool_true_stable => {
        r#"void main() {
  var flag = true;
  print(flag.hashCode == flag.hashCode);
}"#,
        ["true"]
    };

    hashcode_list_same_reference_stable => {
        r#"void main() {
  var list = [1, 2, 3];
  print(list.hashCode == list.hashCode);
}"#,
        ["true"]
    };

    hashcode_map_same_reference_stable => {
        r#"void main() {
  var map = {'a': 1};
  print(map.hashCode == map.hashCode);
}"#,
        ["true"]
    };

    hashcode_custom_object_same_instance_stable => {
        r#"class Widget {
  int id;
  Widget(this.id);
}
void main() {
  var w = Widget(5);
  print(w.hashCode == w.hashCode);
}"#,
        ["true"]
    };

    hashcode_equal_strings_share_hash => {
        r#"void main() {
  var a = 'dart';
  var b = 'dart';
  print(a == b);
  print(a.hashCode == b.hashCode);
}"#,
        ["true", "true"]
    };

    hashcode_equal_ints_share_hash => {
        r#"void main() {
  var a = 7;
  var b = 7;
  print(a.hashCode == b.hashCode);
}"#,
        ["true"]
    };

    hashcode_custom_equal_objects_share_hash => {
        r#"class Point {
  int x;
  int y;
  Point(this.x, this.y);
  bool operator ==(Object other) {
    if (other is Point) {
      return x == other.x && y == other.y;
    }
    return false;
  }
  int get hashCode => x * 31 + y;
}
void main() {
  var a = Point(2, 3);
  var b = Point(2, 3);
  print(a == b);
  print(a.hashCode == b.hashCode);
}"#,
        ["true", "true"]
    };

    hashcode_custom_unequal_objects_differ => {
        r#"class Point {
  int x;
  Point(this.x);
  bool operator ==(Object other) {
    if (other is Point) {
      return x == other.x;
    }
    return false;
  }
  int get hashCode => x;
}
void main() {
  var a = Point(1);
  var b = Point(2);
  print(a.hashCode == b.hashCode);
}"#,
        ["false"]
    };

    hashcode_list_after_mutation_stays_stable => {
        r#"void main() {
  var list = [1];
  var before = list.hashCode;
  list.add(2);
  print(list.hashCode == before);
}"#,
        ["true"]
    };

    hashcode_equal_lists_share_hash => {
        r#"void main() {
  var a = [1, 2];
  var b = [1, 2];
  print(a == b);
  print(a.hashCode == b.hashCode);
}"#,
        ["true", "true"]
    };

    hashcode_unequal_lists_differ => {
        r#"void main() {
  var a = [1, 2];
  var b = [3, 4];
  print(a == b);
  print(a.hashCode == b.hashCode);
}"#,
        ["false", "false"]
    };

    hashcode_enum_member_stable => {
        r#"enum Color { red, green, blue }
void main() {
  var c = Color.green;
  print(c.hashCode == c.hashCode);
}"#,
        ["true"]
    };

    hashcode_double_stable => {
        r#"void main() {
  var x = 3.14;
  print(x.hashCode == x.hashCode);
}"#,
        ["true"]
    };

    hashcode_empty_string_stable => {
        r#"void main() {
  var s = '';
  print(s.hashCode == s.hashCode);
}"#,
        ["true"]
    };

    hashcode_zero_int_stable => {
        r#"void main() {
  var n = 0;
  print(n.hashCode == n.hashCode);
}"#,
        ["true"]
    };

    hashcode_set_same_reference_stable => {
        r#"void main() {
  var set = {1, 2, 3};
  print(set.hashCode == set.hashCode);
}"#,
        ["true"]
    };

    hashcode_record_same_reference_stable => {
        r#"void main() {
  var pair = (1, 'x');
  print(pair.hashCode == pair.hashCode);
}"#,
        ["true"]
    };

    hashcode_custom_default_differs_for_two_instances => {
        r#"class Box {
  int v;
  Box(this.v);
}
void main() {
  var a = Box(1);
  var b = Box(1);
  print(a == b);
  print(a.hashCode == b.hashCode);
}"#,
        ["false", "false"]
    };
}
