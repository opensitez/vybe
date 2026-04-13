use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// ECMAScript: Error handling — try/catch/finally, throw,
// custom errors, error types
// ═══════════════════════════════════════════════════════════

#[test]
fn try_catch_basic() {
    let out = run_js(r#"
try {
    throw new Error("oops");
} catch (e) {
    console.log(e.message);
}
"#);
    assert_eq!(out, vec!["oops"]);
}

#[test]
fn try_catch_no_error() {
    let out = run_js(r#"
try {
    console.log("ok");
} catch (e) {
    console.log("error");
}
"#);
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn try_finally() {
    let out = run_js(r#"
try {
    console.log("try");
} finally {
    console.log("finally");
}
"#);
    assert_eq!(out, vec!["try", "finally"]);
}

#[test]
fn try_catch_finally() {
    let out = run_js(r#"
try {
    throw new Error("fail");
} catch (e) {
    console.log("caught: " + e.message);
} finally {
    console.log("cleanup");
}
"#);
    assert_eq!(out, vec!["caught: fail", "cleanup"]);
}

#[ignore]
#[test]
fn finally_always_runs() {
    let out = run_js(r#"
function test() {
    try {
        return 1;
    } finally {
        console.log("finally");
    }
}
test();
"#);
    assert_eq!(out, vec!["finally"]);
}

#[test]
fn throw_string() {
    let out = run_js(r#"
try {
    throw "simple error";
} catch (e) {
    console.log(e);
}
"#);
    assert_eq!(out, vec!["simple error"]);
}

#[test]
fn throw_number() {
    let out = run_js(r#"
try {
    throw 404;
} catch (e) {
    console.log(e);
}
"#);
    assert_eq!(out, vec!["404"]);
}

#[test]
fn error_with_message() {
    let out = run_js(r#"
try {
    throw new Error("something went wrong");
} catch (e) {
    console.log(e.message);
}
"#);
    assert_eq!(out, vec!["something went wrong"]);
}

#[test]
fn nested_try_catch() {
    let out = run_js(r#"
try {
    try {
        throw new Error("inner");
    } catch (e) {
        console.log("inner: " + e.message);
        throw new Error("rethrown");
    }
} catch (e) {
    console.log("outer: " + e.message);
}
"#);
    assert_eq!(out, vec!["inner: inner", "outer: rethrown"]);
}

#[ignore]
#[test]
fn custom_error_class() {
    let out = run_js(r#"
class ValidationError extends Error {
    constructor(message, field) {
        super(message);
        this.field = field;
    }
}
try {
    throw new ValidationError("Required", "name");
} catch (e) {
    console.log(e.message);
    console.log(e.field);
}
"#);
    assert_eq!(out, vec!["Required", "name"]);
}

#[test]
fn error_in_function() {
    let out = run_js(r#"
function divide(a, b) {
    if (b === 0) throw new Error("Division by zero");
    return a / b;
}
try {
    console.log(divide(10, 2));
    console.log(divide(10, 0));
} catch (e) {
    console.log(e.message);
}
"#);
    assert_eq!(out, vec!["5", "Division by zero"]);
}

#[test]
fn catch_without_binding() {
    let out = run_js(r#"
try {
    throw new Error("ignored");
} catch {
    console.log("caught");
}
"#);
    assert_eq!(out, vec!["caught"]);
}

#[ignore]
#[test]
fn error_instanceof() {
    let out = run_js(r#"
class AppError extends Error {
    constructor(msg) { super(msg); }
}
try {
    throw new AppError("test");
} catch (e) {
    console.log(e instanceof AppError);
    console.log(e instanceof Error);
}
"#);
    assert_eq!(out, vec!["true", "true"]);
}
