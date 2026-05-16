use super::helpers::{compile_ok, run_prints};

// ── Custom exception classes ─────────────────────────────────

#[test] fn custom_exception() {
    compile_ok("class AppException implements Exception { final String message; AppException(this.message); }");
}

#[test] fn custom_exception_thrown() {
    compile_ok(r#"
class AppException implements Exception {
  final String message;
  AppException(this.message);
}
void risky() { throw AppException('Something went wrong'); }
"#);
}

#[test] fn custom_exception_caught() {
    let out = run_prints(r#"
class AppException implements Exception {
  final String message;
  AppException(this.message);
}
void main() {
  try {
    throw AppException('oops');
  } catch (e) {
    print('caught');
  }
}
"#);
    assert_eq!(out, ["caught"]);
}

#[test] fn custom_error() {
    compile_ok("class StateError extends Error { final String msg; StateError(this.msg); }");
}

// ── on — typed catch ─────────────────────────────────────────

#[test] fn on_exception_type() {
    compile_ok(r#"
void main() {
  try {
    throw FormatException('bad format');
  } on FormatException catch (e) {
    print('format error');
  }
}
"#);
}

#[test] fn on_without_var() {
    compile_ok(r#"
void main() {
  try {
    throw FormatException('bad');
  } on FormatException {
    print('caught format');
  }
}
"#);
}

#[test] fn on_multiple() {
    compile_ok(r#"
void main() {
  try {
    throw RangeError('out of range');
  } on FormatException {
    print('format');
  } on RangeError catch (e) {
    print('range');
  }
}
"#);
}

#[test] fn on_with_catch_fallback() {
    compile_ok(r#"
void main() {
  try {
    throw Exception('any');
  } on FormatException {
    print('format');
  } catch (e) {
    print('fallback');
  }
}
"#);
}

// ── finally ──────────────────────────────────────────────────

#[test] fn finally_always_runs() {
    let out = run_prints(r#"
void main() {
  try {
    print('try');
  } finally {
    print('finally');
  }
}
"#);
    assert_eq!(out, ["try", "finally"]);
}

#[test] fn finally_after_exception() {
    let out = run_prints(r#"
void main() {
  try {
    throw 'error';
  } catch (e) {
    print('caught');
  } finally {
    print('cleanup');
  }
}
"#);
    assert_eq!(out, ["caught", "cleanup"]);
}

#[test] fn finally_no_catch() {
    compile_ok(r#"
void risky() {
  try {
    var x = 1;
  } finally {
    print('done');
  }
}
"#);
}

// ── rethrow ──────────────────────────────────────────────────

#[test] fn rethrow_basic() {
    compile_ok(r#"
void handle() {
  try {
    throw 'inner error';
  } catch (e) {
    rethrow;
  }
}
"#);
}

#[test] fn rethrow_wrapped() {
    compile_ok(r#"
void inner() { throw 'something'; }
void outer() {
  try {
    inner();
  } catch (e) {
    rethrow;
  }
}
"#);
}

// ── Throwing different types ─────────────────────────────────

#[test] fn throw_string() { compile_ok("void main() { try { throw 'an error'; } catch (e) { print(e); } }"); }
#[test] fn throw_exception_class() { compile_ok("void main() { try { throw Exception('msg'); } catch (e) { } }"); }
#[test] fn throw_in_function() { compile_ok("void fail(String msg) { throw Exception(msg); }"); }

#[test] fn throw_string_result() {
    let out = run_prints("void main() { try { throw 'bad'; } catch (e) { print('got: $e'); } }");
    assert_eq!(out, ["got: bad"]);
}

// ── Stack trace ──────────────────────────────────────────────

#[test] fn catch_with_stack_trace() {
    compile_ok(r#"
void main() {
  try {
    throw Exception('boom');
  } catch (e, s) {
    print('caught');
  }
}
"#);
}

// ── Nested try/catch ─────────────────────────────────────────

#[test] fn nested_try() {
    let out = run_prints(r#"
void main() {
  try {
    try {
      throw 'inner';
    } catch (e) {
      print('inner caught: $e');
    }
    print('outer ok');
  } catch (e) {
    print('outer caught');
  }
}
"#);
    assert_eq!(out, ["inner caught: inner", "outer ok"]);
}

// ── Exceptions in methods ────────────────────────────────────

#[test] fn exception_in_method() {
    compile_ok(r#"
class Parser {
  int parse(String s) {
    if (s.isEmpty) throw FormatException('empty');
    return int.parse(s);
  }
}
"#);
}

#[test] fn exception_propagation() {
    let out = run_prints(r#"
void risky() { throw 'error'; }
void safe() {
  try { risky(); } catch (e) { print('safe caught: $e'); }
}
void main() { safe(); }
"#);
    assert_eq!(out, ["safe caught: error"]);
}

// ── Assert ───────────────────────────────────────────────────

#[test] fn assert_basic() { compile_ok("void main() { assert(1 + 1 == 2); }"); }
#[test] fn assert_with_message() { compile_ok("void main() { assert(true, 'must be true'); }"); }
