use super::helpers::{compile_ok, run_prints};

// ── is — type checks ────────────────────────────────────────

#[test] fn is_string() { compile_ok("var x = 'hello'; var b = x is String;"); }
#[test] fn is_int() { compile_ok("var x = 42; var b = x is int;"); }
#[test] fn is_double() { compile_ok("var x = 3.14; var b = x is double;"); }
#[test] fn is_bool() { compile_ok("var x = true; var b = x is bool;"); }
#[test] fn is_list() { compile_ok("var x = [1, 2]; var b = x is List;"); }
#[test] fn is_map() { compile_ok("var x = {'a': 1}; var b = x is Map;"); }

#[test] fn is_class() {
    compile_ok("class Dog {} void main() { var d = Dog(); var b = d is Dog; }");
}

#[test] fn is_parent_class() {
    compile_ok("class Animal {} class Dog extends Animal {} void main() { var d = Dog(); var b = d is Animal; }");
}

#[test] fn is_result() {
    let out = run_prints("void main() { var x = 42; print(x is int); }");
    assert_eq!(out, ["true"]);
}

#[test] fn is_string_result() {
    let out = run_prints("void main() { var x = 'hello'; print(x is String); }");
    assert_eq!(out, ["true"]);
}

// ── is! — negated type checks ────────────────────────────────

#[test] fn is_not_string() { compile_ok("var x = 42; var b = x is! String;"); }
#[test] fn is_not_int() { compile_ok("var x = 'hello'; var b = x is! int;"); }

#[test] fn is_not_result() {
    let out = run_prints("void main() { var x = 'hello'; print(x is! int); }");
    assert_eq!(out, ["true"]);
}

// ── as — type casts ─────────────────────────────────────────

#[test] fn as_num() { compile_ok("dynamic x = 42; var n = x as int;"); }
#[test] fn as_string() { compile_ok("dynamic x = 'hello'; var s = x as String;"); }

#[test] fn as_parent() {
    compile_ok("class Animal { String name = 'A'; } class Dog extends Animal {} void main() { Dog d = Dog(); Animal a = d as Animal; }");
}

// ── dynamic type ────────────────────────────────────────────

#[test] fn dynamic_var() { compile_ok("dynamic x = 42; x = 'hello'; x = true;"); }
#[test] fn dynamic_in_function() { compile_ok("dynamic f(dynamic x) { return x; }"); }
#[test] fn dynamic_field() { compile_ok("class Bag { dynamic content; Bag(this.content); }"); }

#[test] fn dynamic_result() {
    let out = run_prints("void main() { dynamic x = 42; print(x); }");
    assert_eq!(out, ["42"]);
}

// ── Object type ──────────────────────────────────────────────

#[test] fn object_param() { compile_ok("void log(Object value) { print(value); }"); }
#[test] fn object_is_parent() {
    compile_ok("class Foo {} void main() { var f = Foo(); print(f is Object); }");
}

// ── Nullable annotations ─────────────────────────────────────

#[test] fn nullable_string() { compile_ok("String? name;"); }
#[test] fn nullable_int() { compile_ok("int? count;"); }
#[test] fn nullable_param() { compile_ok("void greet(String? name) { print(name ?? 'guest'); }"); }
#[test] fn nullable_return() { compile_ok("String? findName(bool found) { return found ? 'Alice' : null; }"); }

#[test] fn nullable_field() {
    compile_ok("class Node { int value; Node? next; Node(this.value); }");
}

#[test] fn nullable_chain() {
    compile_ok("class A { int x = 1; } void main() { A? a = null; var v = a?.x; }");
}

#[test] fn nullable_coalesce() {
    let out = run_prints("void main() { String? s = null; print(s ?? 'default'); }");
    assert_eq!(out, ["default"]);
}

// ── ! non-null assertion ─────────────────────────────────────

#[test] fn non_null_assert() { compile_ok("String? s = 'hello'; var n = s!;"); }
#[test] fn non_null_in_method() { compile_ok("class A { int? x = 5; int get val => x!; }"); }

#[test] fn non_null_result() {
    let out = run_prints("void main() { String? s = 'world'; print(s!); }");
    assert_eq!(out, ["world"]);
}

// ── Type narrowing with is ───────────────────────────────────

#[test] fn is_in_if() {
    compile_ok("void f(dynamic x) { if (x is String) { print(x.toUpperCase()); } }");
}

#[test] fn is_in_switch() {
    compile_ok(r#"
void describe(dynamic x) {
  if (x is int) {
    print('int: $x');
  } else if (x is String) {
    print('string: $x');
  } else {
    print('other');
  }
}
"#);
}

// ── runtimeType ─────────────────────────────────────────────

#[test] fn runtime_type_int() { compile_ok("var x = 42; var t = x.runtimeType;"); }
#[test] fn runtime_type_string() { compile_ok("var x = 'hi'; var t = x.runtimeType;"); }

// ── Type inference ──────────────────────────────────────────

#[test] fn var_inference_int() {
    let out = run_prints("void main() { var x = 5; print(x); }");
    assert_eq!(out, ["5"]);
}

#[test] fn var_inference_bool() {
    let out = run_prints("void main() { var b = 5 > 3; print(b); }");
    assert_eq!(out, ["true"]);
}

#[test] fn inferred_list_type() { compile_ok("var nums = [1, 2, 3]; var first = nums.first;"); }
#[test] fn inferred_map_type() { compile_ok("var m = {'key': 42}; var v = m['key'];"); }
