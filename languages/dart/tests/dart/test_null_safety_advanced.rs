use super::helpers::{compile_ok, run_prints};

// ── Nullable field patterns ──────────────────────────────────

#[test]
fn nullable_field_class() {
    compile_ok("class Person { String? middleName; String first; Person(this.first); }");
}

#[test]
fn nullable_field_default() {
    let out = run_prints(
        r#"
class Person { String? nickname; String name; Person(this.name); }
void main() { var p = Person('Alice'); print(p.nickname ?? 'no nickname'); }
"#,
    );
    assert_eq!(out, ["no nickname"]);
}

#[test]
fn nullable_chain_two_levels() {
    compile_ok(
        "class B { int x = 1; } class A { B? b; } void main() { var a = A(); var v = a.b?.x; }",
    );
}

#[test]
fn nullable_chain_method() {
    compile_ok(
        "class A { String? name; } void main() { var a = A(); var v = a.name?.toUpperCase(); }",
    );
}

#[test]
fn nullable_chain_result() {
    let out = run_prints(
        r#"
class A { String? name = 'alice'; }
void main() { var a = A(); print(a.name?.toUpperCase()); }
"#,
    );
    assert_eq!(out, ["ALICE"]);
}

#[test]
fn nullable_null_chain_result() {
    let out = run_prints(
        r#"
class A { String? name; }
void main() { var a = A(); var r = a.name?.toUpperCase() ?? 'null'; print(r); }
"#,
    );
    assert_eq!(out, ["null"]);
}

// ── ?? coalescing patterns ───────────────────────────────────

#[test]
fn coalesce_null() {
    let out = run_prints("void main() { var x = null; print(x ?? 'fallback'); }");
    assert_eq!(out, ["fallback"]);
}

#[test]
fn coalesce_not_null() {
    let out = run_prints("void main() { var x = 'real'; print(x ?? 'fallback'); }");
    assert_eq!(out, ["real"]);
}

#[test]
fn coalesce_chained() {
    let out =
        run_prints("void main() { String? a; String? b; var c = a ?? b ?? 'default'; print(c); }");
    assert_eq!(out, ["default"]);
}

#[test]
fn coalesce_with_expr() {
    let out = run_prints("void main() { int? n; print((n ?? 0) + 10); }");
    assert_eq!(out, ["10"]);
}

// ── ??= assignment ────────────────────────────────────────────

#[test]
fn null_assign_basic() {
    let out = run_prints("void main() { String? s; s ??= 'hello'; print(s); }");
    assert_eq!(out, ["hello"]);
}

#[test]
fn null_assign_no_overwrite() {
    let out = run_prints("void main() { String? s = 'existing'; s ??= 'new'; print(s); }");
    assert_eq!(out, ["existing"]);
}

#[test]
fn null_assign_field() {
    compile_ok(
        r#"
class Cache {
  String? _value;
  String get value { _value ??= 'default'; return _value!; }
}
"#,
    );
}

// ── ! assertion ──────────────────────────────────────────────

#[test]
fn assert_non_null() {
    let out = run_prints("void main() { String? s = 'hello'; print(s!.toUpperCase()); }");
    assert_eq!(out, ["HELLO"]);
}

#[test]
fn assert_non_null_int() {
    let out = run_prints("void main() { int? n = 42; print(n! + 1); }");
    assert_eq!(out, ["43"]);
}

#[test]
fn assert_non_null_method() {
    compile_ok("class A { int? val = 5; void f() { print(val!); } }");
}

// ── Nullable function params ─────────────────────────────────

#[test]
fn nullable_param_with_default() {
    let out = run_prints(
        r#"
void greet(String? name) { print('Hello ${name ?? 'stranger'}'); }
void main() { greet(null); }
"#,
    );
    assert_eq!(out, ["Hello stranger"]);
}

#[test]
fn nullable_param_provided() {
    let out = run_prints(
        r#"
void greet(String? name) { print('Hello ${name ?? 'stranger'}'); }
void main() { greet('Alice'); }
"#,
    );
    assert_eq!(out, ["Hello Alice"]);
}

// ── Nullable return types ────────────────────────────────────

#[test]
fn nullable_return_null() {
    let out = run_prints(
        r#"
String? find(List<String> list, String key) {
  for (var s in list) { if (s == key) return s; }
  return null;
}
void main() { var r = find(['a', 'b'], 'c'); print(r ?? 'not found'); }
"#,
    );
    assert_eq!(out, ["not found"]);
}

#[test]
fn nullable_return_found() {
    let out = run_prints(
        r#"
String? find(List<String> list, String key) {
  for (var s in list) { if (s == key) return s; }
  return null;
}
void main() { var r = find(['a', 'b'], 'a'); print(r ?? 'missing'); }
"#,
    );
    assert_eq!(out, ["a"]);
}

// ── Null safety with collections ─────────────────────────────

#[test]
fn nullable_list_element() {
    compile_ok("List<String?> list = ['a', null, 'b'];");
}

#[test]
fn nullable_map_value() {
    compile_ok("Map<String, int?> m = {'a': 1, 'b': null};");
}

#[test]
fn nullable_list_filter() {
    compile_ok(
        "List<String?> list = ['a', null, 'b']; var nonNull = list.where((e) => e != null).toList();",
    );
}

// ── Late + nullable combos ───────────────────────────────────

#[test]
fn late_nullable() {
    compile_ok("late String? name; void main() { name = null; print(name ?? 'empty'); }");
}

#[test]
fn late_initialized_with_nullable() {
    compile_ok(
        r#"
class Loader {
  late String? data;
  void load(bool found) { data = found ? 'result' : null; }
}
"#,
    );
}

// ── Null checks in control flow ──────────────────────────────

#[test]
fn null_check_if() {
    let out = run_prints(
        r#"
void main() {
  String? s = 'hello';
  if (s != null) { print(s.toUpperCase()); }
}
"#,
    );
    assert_eq!(out, ["HELLO"]);
}

#[test]
fn null_guard_return() {
    compile_ok(
        r#"
void process(String? s) {
  if (s == null) return;
  print(s.toUpperCase());
}
"#,
    );
}

#[test]
fn null_check_ternary() {
    let out = run_prints(
        r#"
void main() {
  String? name = null;
  print(name != null ? name : 'anonymous');
}
"#,
    );
    assert_eq!(out, ["anonymous"]);
}
