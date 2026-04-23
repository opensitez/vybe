use super::helpers::{compile_ok, run_prints};

// ── final variables ─────────────────────────────────────────

#[test] fn final_var() { compile_ok("final x = 42;"); }
#[test] fn final_typed() { compile_ok("final int x = 42;"); }
#[test] fn final_string() { compile_ok("final name = 'Dart';"); }
#[test] fn final_list() { compile_ok("final items = [1, 2, 3];"); }
#[test] fn final_map() { compile_ok("final m = {'a': 1};"); }
#[test] fn final_in_function() { compile_ok("void f() { final x = 10; print(x); }"); }

#[test] fn final_class_field() {
    compile_ok("class Point { final int x; final int y; Point(this.x, this.y); }");
}

#[test] fn final_class_field_initialized() {
    compile_ok("class Config { final String host = 'localhost'; }");
}

#[test] fn final_in_loop() {
    compile_ok("void main() { for (var i = 0; i < 3; i++) { final v = i * 2; print(v); } }");
}

#[test] fn final_result() {
    let out = run_prints("void main() { final x = 7; print(x); }");
    assert_eq!(out, ["7"]);
}

// ── const values ────────────────────────────────────────────

#[test] fn const_int() { compile_ok("const x = 42;"); }
#[test] fn const_string() { compile_ok("const greeting = 'Hello';"); }
#[test] fn const_double() { compile_ok("const pi = 3.14159;"); }
#[test] fn const_bool() { compile_ok("const debug = true;"); }
#[test] fn const_typed_int() { compile_ok("const int max = 100;"); }
#[test] fn const_typed_string() { compile_ok("const String prefix = 'ID-';"); }

#[test] fn const_in_class() {
    compile_ok("class Config { static const int timeout = 30; static const String host = 'localhost'; }");
}

#[test] fn const_used_in_expr() {
    compile_ok("const base = 10; var doubled = base * 2;");
}

#[test] fn const_result() {
    let out = run_prints("void main() { const x = 99; print(x); }");
    assert_eq!(out, ["99"]);
}

#[test] fn const_top_level() {
    compile_ok("const maxRetries = 3; const baseUrl = 'https://api.example.com';");
}

#[test] fn const_list() { compile_ok("const items = [1, 2, 3];"); }
#[test] fn const_map() { compile_ok("const m = {'key': 'value'};"); }

#[test] fn const_constructor() {
    compile_ok("class Color { final int r; final int g; final int b; const Color(this.r, this.g, this.b); } const red = Color(255, 0, 0);");
}

#[test] fn const_class_used() {
    compile_ok("class Size { final int w; final int h; const Size(this.w, this.h); } void main() { const s = Size(100, 200); print(s.w); }");
}

// ── late variables ──────────────────────────────────────────

#[test] fn late_var() { compile_ok("late int x; void main() { x = 42; print(x); }"); }
#[test] fn late_typed_string() { compile_ok("late String name; void main() { name = 'Dart'; }"); }

#[test] fn late_in_class() {
    compile_ok("class Lazy { late int value; void init() { value = 100; } }");
}

#[test] fn late_final() {
    compile_ok("class Once { late final int id; void setId(int v) { id = v; } }");
}

#[test] fn late_initialized_field() {
    compile_ok("class Db { late String connection; Db() { connection = 'sqlite:memory'; } String get conn => connection; }");
}

#[test] fn late_result() {
    let out = run_prints("late int x; void main() { x = 55; print(x); }");
    assert_eq!(out, ["55"]);
}

#[test] fn late_final_result() {
    let out = run_prints(r#"
class Counter {
  late final int start;
  Counter(int v) { start = v; }
}
void main() {
  var c = Counter(10);
  print(c.start);
}
"#);
    assert_eq!(out, ["10"]);
}

// ── combinations ────────────────────────────────────────────

#[test] fn final_and_const_together() {
    compile_ok("const limit = 100; final result = limit * 2;");
}

#[test] fn final_late_const_in_class() {
    compile_ok(r#"
class Settings {
  static const String version = '1.0';
  final String name;
  late String description;
  Settings(this.name);
}
"#);
}

#[test] fn const_in_switch() {
    compile_ok(r#"
const kMax = 3;
void main() {
  switch (kMax) {
    case 3: print('three'); break;
    default: print('other');
  }
}
"#);
}

#[test] fn final_computed() {
    let out = run_prints("void main() { final a = 6; final b = 7; final c = a * b; print(c); }");
    assert_eq!(out, ["42"]);
}

#[test] fn const_math_expr() {
    compile_ok("const half = 1 / 2; const double_pi = 3.14 * 2;");
}
