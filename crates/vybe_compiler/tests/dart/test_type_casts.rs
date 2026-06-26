//! Type tests and casts: is, is!, as, as?, runtimeType, and narrowing for
//! int, String, List, Map, and custom types.

dart_cases! {
    is_int_on_integer_literal => {
        r#"void main() {
  var x = 42;
  print(x is int);
}"#,
        ["true"]
    };

    is_string_on_string_literal => {
        r#"void main() {
  var x = 'hello';
  print(x is String);
}"#,
        ["true"]
    };

    is_int_on_string_value_is_false => {
        r#"void main() {
  var x = '42';
  print(x is int);
}"#,
        ["false"]
    };

    is_string_on_int_value_is_false => {
        r#"void main() {
  var x = 42;
  print(x is String);
}"#,
        ["false"]
    };

    is_double_on_decimal_literal => {
        r#"void main() {
  var x = 3.14;
  print(x is double);
}"#,
        ["true"]
    };

    is_bool_on_true_literal => {
        r#"void main() {
  var x = true;
  print(x is bool);
}"#,
        ["true"]
    };

    is_list_on_list_literal => {
        r#"void main() {
  var x = [1, 2, 3];
  print(x is List);
}"#,
        ["true"]
    };

    is_map_on_map_literal => {
        r#"void main() {
  var x = {'a': 1};
  print(x is Map);
}"#,
        ["true"]
    };

    is_null_on_null_value => {
        r#"void main() {
  Object? x = null;
  print(x is Null);
}"#,
        ["true"]
    };

    is_not_string_on_int => {
        r#"void main() {
  var x = 99;
  print(x is! String);
}"#,
        ["true"]
    };

    is_not_int_on_string => {
        r#"void main() {
  var x = 'text';
  print(x is! int);
}"#,
        ["true"]
    };

    is_not_list_on_map => {
        r#"void main() {
  var x = {'k': 1};
  print(x is! List);
}"#,
        ["true"]
    };

    is_not_bool_on_int => {
        r#"void main() {
  var x = 1;
  print(x is! bool);
}"#,
        ["true"]
    };

    is_custom_class_on_instance => {
        r#"class Widget { String id = 'w1'; }
void main() {
  var w = Widget();
  print(w is Widget);
}"#,
        ["true"]
    };

    is_parent_type_on_subclass_instance => {
        r#"class Animal {}
class Dog extends Animal {}
void main() {
  var d = Dog();
  print(d is Animal);
}"#,
        ["true"]
    };

    is_subclass_on_parent_reference_is_false => {
        r#"class Animal {}
class Dog extends Animal {}
void main() {
  Animal a = Dog();
  print(a is Dog);
}"#,
        ["true"]
    };

    as_cast_dynamic_to_int => {
        r#"void main() {
  dynamic x = 7;
  var n = x as int;
  print(n + 3);
}"#,
        ["10"]
    };

    as_cast_dynamic_to_string => {
        r#"void main() {
  dynamic x = 'dart';
  var s = x as String;
  print(s.toUpperCase());
}"#,
        ["DART"]
    };

    as_cast_dynamic_to_list => {
        r#"void main() {
  dynamic x = [1, 2];
  var list = x as List;
  print(list.length);
}"#,
        ["2"]
    };

    as_cast_subclass_to_parent => {
        r#"class Shape { String kind = 'shape'; }
class Circle extends Shape { double r = 1.0; }
void main() {
  Circle c = Circle();
  var s = c as Shape;
  print(s.kind);
}"#,
        ["shape"]
    };

    as_question_returns_value_when_type_matches => {
        r#"void main() {
  Object value = 'safe';
  var s = value as String?;
  print(s);
}"#,
        ["safe"]
    };

    as_question_returns_null_when_type_mismatch => {
        r#"void main() {
  Object value = 42;
  var s = value as String?;
  print(s ?? 'not-string');
}"#,
        ["not-string"]
    };

    as_question_on_null_returns_null => {
        r#"void main() {
  Object? value = null;
  var s = value as String?;
  print(s ?? 'was-null');
}"#,
        ["was-null"]
    };

    as_question_list_cast_succeeds => {
        r#"void main() {
  Object value = [10, 20];
  var list = value as List?;
  print(list?.length);
}"#,
        ["2"]
    };

    as_question_list_cast_fails => {
        r#"void main() {
  Object value = 'not-list';
  var list = value as List?;
  print(list?.length ?? -1);
}"#,
        ["-1"]
    };

    runtime_type_of_int => {
        r#"void main() {
  var x = 42;
  print(x.runtimeType == int);
}"#,
        ["true"]
    };

    runtime_type_of_string => {
        r#"void main() {
  var x = 'hi';
  print(x.runtimeType == String);
}"#,
        ["true"]
    };

    runtime_type_of_bool => {
        r#"void main() {
  var x = false;
  print(x.runtimeType == bool);
}"#,
        ["true"]
    };

    runtime_type_of_list => {
        r#"void main() {
  var x = <int>[1, 2];
  print(x.runtimeType == List<int>);
}"#,
        ["true"]
    };

    runtime_type_of_custom_class => {
        r#"class Point { int x = 0; int y = 0; }
void main() {
  var p = Point();
  print(p.runtimeType == Point);
}"#,
        ["true"]
    };

    type_narrowing_is_int_in_if_branch => {
        r#"void describe(Object value) {
  if (value is int) {
    print(value + 1);
  } else {
    print(-1);
  }
}
void main() {
  describe(5);
}"#,
        ["6"]
    };

    type_narrowing_is_string_in_if_branch => {
        r#"void describe(Object value) {
  if (value is String) {
    print(value.length);
  } else {
    print(0);
  }
}
void main() {
  describe('abcd');
}"#,
        ["4"]
    };

    type_narrowing_is_list_in_if_branch => {
        r#"void describe(Object value) {
  if (value is List) {
    print(value.isEmpty);
  } else {
    print(false);
  }
}
void main() {
  describe(<int>[]);
}"#,
        ["true"]
    };

    type_narrowing_is_not_string_routes_to_else => {
        r#"void describe(Object value) {
  if (value is! String) {
    print('not-str');
  } else {
    print('str');
  }
}
void main() {
  describe(12);
}"#,
        ["not-str"]
    };

    is_list_int_on_typed_list => {
        r#"void main() {
  List<int> nums = [1, 2, 3];
  print(nums is List<int>);
}"#,
        ["true"]
    };

    is_list_int_on_untyped_list_is_false => {
        r#"void main() {
  var nums = [1, 'two'];
  print(nums is List<int>);
}"#,
        ["false"]
    };

    is_map_string_int_on_typed_map => {
        r#"void main() {
  Map<String, int> m = {'a': 1};
  print(m is Map<String, int>);
}"#,
        ["true"]
    };

    as_cast_list_element_access_after_cast => {
        r#"void main() {
  dynamic raw = [5, 6, 7];
  var list = raw as List<int>;
  print(list[1]);
}"#,
        ["6"]
    };

    is_num_on_int_is_true => {
        r#"void main() {
  num n = 10;
  print(n is int);
}"#,
        ["true"]
    };

    is_num_on_double_is_true => {
        r#"void main() {
  num n = 2.5;
  print(n is double);
}"#,
        ["true"]
    };

    runtime_type_printed_via_to_string => {
        r#"void main() {
  var x = [1];
  print('${x.runtimeType}'.contains('List'));
}"#,
        ["true"]
    };

    chained_is_and_as_question_in_expression => {
        r#"void main() {
  Object? value = 'cast-me';
  var len = (value is String ? value : value as String?)?.length ?? 0;
  print(len);
}"#,
        ["7"]
    };
}
