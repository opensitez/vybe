//! Default Object protocol: hashCode, toString, runtimeType; identical vs ==.

dart_cases! {
    int_to_string_default => {
        r#"void main() {
  print(42.toString());
}"#,
        ["42"]
    };

    int_runtime_type_is_int => {
        r#"void main() {
  var n = 42;
  print(n.runtimeType == int);
}"#,
        ["true"]
    };

    int_hashcode_equals_itself => {
        r#"void main() {
  var n = 99;
  print(n.hashCode == n.hashCode);
}"#,
        ["true"]
    };

    int_identical_vs_equal_literals => {
        r#"void main() {
  print(identical(7, 7));
  print(7 == 7);
}"#,
        ["true", "true"]
    };

    string_to_string_returns_self_content => {
        r#"void main() {
  print('hello'.toString());
}"#,
        ["hello"]
    };

    string_runtime_type_is_string => {
        r#"void main() {
  var s = 'dart';
  print(s.runtimeType == String);
}"#,
        ["true"]
    };

    string_hashcode_stable => {
        r#"void main() {
  var s = 'abc';
  print(s.hashCode == s.hashCode);
}"#,
        ["true"]
    };

    string_identical_vs_equal_different_instances => {
        r#"void main() {
  var a = String.fromCharCode(65);
  var b = String.fromCharCode(65);
  print(identical(a, b));
  print(a == b);
}"#,
        ["false", "true"]
    };

    list_to_string_default_format => {
        r#"void main() {
  print([1, 2, 3].toString());
}"#,
        ["[1, 2, 3]"]
    };

    list_runtime_type_is_list => {
        r#"void main() {
  var list = [1, 2];
  print(list.runtimeType == List);
}"#,
        ["true"]
    };

    list_hashcode_stable_for_same_reference => {
        r#"void main() {
  var list = [1, 2];
  print(list.hashCode == list.hashCode);
}"#,
        ["true"]
    };

    list_identical_false_equal_true_for_same_content => {
        r#"void main() {
  print(identical([1], [1]));
  print([1] == [1]);
}"#,
        ["false", "true"]
    };

    list_identical_true_for_same_variable => {
        r#"void main() {
  var a = [1, 2];
  print(identical(a, a));
}"#,
        ["true"]
    };

    double_to_string_default => {
        r#"void main() {
  print(3.5.toString());
}"#,
        ["3.5"]
    };

    double_runtime_type_is_double => {
        r#"void main() {
  var x = 2.5;
  print(x.runtimeType == double);
}"#,
        ["true"]
    };

    bool_to_string_true => {
        r#"void main() {
  print(true.toString());
}"#,
        ["true"]
    };

    bool_runtime_type_is_bool => {
        r#"void main() {
  print(false.runtimeType == bool);
}"#,
        ["true"]
    };

    null_runtime_type_is_null => {
        r#"void main() {
  print(null.runtimeType == Null);
}"#,
        ["true"]
    };

    null_to_string_returns_null => {
        r#"void main() {
  print(null.toString());
}"#,
        ["null"]
    };

    class_default_to_string_contains_instance_of => {
        r#"class Widget {}
void main() {
  print(Widget().toString().contains('Widget'));
}"#,
        ["true"]
    };

    class_default_hashcode_differs_between_instances => {
        r#"class Token {
  int id;
  Token(this.id);
}
void main() {
  print(Token(1).hashCode == Token(1).hashCode);
}"#,
        ["false"]
    };

    class_default_equals_is_identity => {
        r#"class Node {
  int v;
  Node(this.v);
}
void main() {
  var a = Node(1);
  var b = Node(1);
  print(a == b);
  print(identical(a, b));
}"#,
        ["false", "false"]
    };

    class_same_instance_equals_and_identical => {
        r#"class Item {
  int id;
  Item(this.id);
}
void main() {
  var x = Item(5);
  print(x == x);
  print(identical(x, x));
}"#,
        ["true", "true"]
    };

    class_runtime_type_matches_declared => {
        r#"class Point {
  int x;
  Point(this.x);
}
void main() {
  print(Point(1).runtimeType == Point);
}"#,
        ["true"]
    };

    custom_to_string_overrides_default => {
        r#"class Labeled {
  String toString() => 'Labeled';
}
void main() {
  print(Labeled().toString());
}"#,
        ["Labeled"]
    };

    custom_hashcode_overrides_default => {
        r#"class Key {
  int k;
  Key(this.k);
  int get hashCode => k;
}
void main() {
  print(Key(3).hashCode);
}"#,
        ["3"]
    };

    map_to_string_default => {
        r#"void main() {
  print({'a': 1, 'b': 2}.toString());
}"#,
        ["{a: 1, b: 2}"]
    };

    map_runtime_type_is_map => {
        r#"void main() {
  var m = {'x': 1};
  print(m.runtimeType == Map);
}"#,
        ["true"]
    };

    map_identical_vs_equal => {
        r#"void main() {
  print(identical({'a': 1}, {'a': 1}));
  print({'a': 1} == {'a': 1});
}"#,
        ["false", "true"]
    };

    set_to_string_default => {
        r#"void main() {
  print({1, 2}.toString().contains('1'));
}"#,
        ["true"]
    };

    set_runtime_type_is_set => {
        r#"void main() {
  var s = {1, 2};
  print(s.runtimeType == Set);
}"#,
        ["true"]
    };

    record_to_string_default => {
        r#"void main() {
  print((1, 'a').toString());
}"#,
        ["(1, a)"]
    };

    record_runtime_type_is_record => {
        r#"void main() {
  var r = (1, 2);
  print(r.runtimeType.toString().contains('Record'));
}"#,
        ["true"]
    };

    enum_to_string_default => {
        r#"enum Color { red, green }
void main() {
  print(Color.red.toString());
}"#,
        ["Color.red"]
    };

    enum_runtime_type_is_enum => {
        r#"enum Mode { on, off }
void main() {
  print(Mode.on.runtimeType == Mode);
}"#,
        ["true"]
    };

    enum_identical_same_member => {
        r#"enum Status { ok, fail }
void main() {
  print(identical(Status.ok, Status.ok));
  print(Status.ok == Status.ok);
}"#,
        ["true", "true"]
    };

    int_double_equal_but_not_identical => {
        r#"void main() {
  print(3 == 3.0);
  print(identical(3, 3.0));
}"#,
        ["true", "false"]
    };

    string_empty_to_string => {
        r#"void main() {
  print(''.toString().isEmpty);
}"#,
        ["true"]
    };

    list_empty_to_string => {
        r#"void main() {
  print([].toString());
}"#,
        ["[]"]
    };

    object_runtime_type_on_int_variable => {
        r#"void main() {
  Object o = 10;
  print(o.runtimeType == int);
}"#,
        ["true"]
    };

    object_runtime_type_on_string_variable => {
        r#"void main() {
  Object o = 'hi';
  print(o.runtimeType == String);
}"#,
        ["true"]
    };

    hashcode_equal_ints_same_value => {
        r#"void main() {
  print(5.hashCode == 5.hashCode);
}"#,
        ["true"]
    };

    hashcode_equal_strings_same_content => {
        r#"void main() {
  print('x'.hashCode == 'x'.hashCode);
}"#,
        ["true"]
    };

    identical_list_alias_equals_true => {
        r#"void main() {
  var a = [1, 2];
  var b = a;
  print(identical(a, b));
  print(a == b);
}"#,
        ["true", "true"]
    };

    class_subclass_runtime_type => {
        r#"class Animal {}
class Dog extends Animal {}
void main() {
  print(Dog().runtimeType == Dog);
}"#,
        ["true"]
    };

    class_supertype_variable_runtime_type => {
        r#"class Base {}
class Child extends Base {}
void main() {
  Base b = Child();
  print(b.runtimeType == Child);
}"#,
        ["true"]
    };

    to_string_on_interpolated_int => {
        r#"void main() {
  var n = 8;
  print('$n');
}"#,
        ["8"]
    };

    runtime_type_to_string_contains_type_name => {
        r#"void main() {
  print(42.runtimeType.toString().contains('int'));
}"#,
        ["true"]
    };

    identical_function_same_reference => {
        r#"void main() {
  void fn() {}
  print(identical(fn, fn));
}"#,
        ["true"]
    };

    equal_custom_objects_with_overridden_equals => {
        r#"class Pair {
  int a;
  int b;
  Pair(this.a, this.b);
  bool operator ==(Object other) {
    return other is Pair && a == other.a && b == other.b;
  }
}
void main() {
  print(Pair(1, 2) == Pair(1, 2));
  print(identical(Pair(1, 2), Pair(1, 2)));
}"#,
        ["true", "false"]
    };

    null_identical_and_equal_to_null => {
        r#"void main() {
  print(null == null);
  print(identical(null, null));
}"#,
        ["true", "true"]
    };
}
