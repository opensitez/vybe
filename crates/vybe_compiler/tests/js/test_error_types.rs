/// JavaScript error types: TypeError, RangeError, ReferenceError, SyntaxError,
/// custom error classes, error properties, error handling patterns,
/// re-throwing, conditional catch, error chaining.

use super::helpers::run_js;

// ===================================================================
// ERROR TYPES
// ===================================================================

#[test]
fn type_error_basic() {
    assert_eq!(run_js(r#"
try {
    null.foo;
} catch (e) {
    console.log(e instanceof TypeError);
    console.log(e.message);
}
"#), &["true", "Cannot read properties of null"]);
}

#[test]
fn range_error() {
    assert_eq!(run_js(r#"
try {
    throw new RangeError("value out of range");
} catch (e) {
    console.log(e instanceof RangeError);
    console.log(e.message);
}
"#), &["true", "value out of range"]);
}

#[test]
fn reference_error() {
    assert_eq!(run_js(r#"
try {
    throw new ReferenceError("x is not defined");
} catch (e) {
    console.log(e instanceof ReferenceError);
    console.log(e.message);
}
"#), &["true", "x is not defined"]);
}

#[test]
fn error_name_property() {
    assert_eq!(run_js(r#"
try {
    throw new TypeError("bad type");
} catch (e) {
    console.log(e.name);
    console.log(e.message);
}
"#), &["TypeError", "bad type"]);
}

#[test]
fn error_instanceof_hierarchy() {
    assert_eq!(run_js(r#"
try {
    throw new TypeError("test");
} catch (e) {
    console.log(e instanceof TypeError);
    console.log(e instanceof Error);
}
"#), &["true", "true"]);
}

// ===================================================================
// CUSTOM ERROR CLASSES
// ===================================================================

#[test]
fn custom_error_class() {
    assert_eq!(run_js(r#"
class AppError extends Error {
    constructor(message, code) {
        super(message);
        this.name = "AppError";
        this.code = code;
    }
}
try {
    throw new AppError("not found", 404);
} catch (e) {
    console.log(e.name);
    console.log(e.message);
    console.log(e.code);
    console.log(e instanceof AppError);
    console.log(e instanceof Error);
}
"#), &["AppError", "not found", "404", "true", "true"]);
}

#[test]
fn custom_error_hierarchy() {
    assert_eq!(run_js(r#"
class HttpError extends Error {
    constructor(status, msg) {
        super(msg);
        this.name = "HttpError";
        this.status = status;
    }
}
class NotFoundError extends HttpError {
    constructor(resource) {
        super(404, resource + " not found");
        this.name = "NotFoundError";
    }
}
try {
    throw new NotFoundError("User");
} catch (e) {
    console.log(e.name);
    console.log(e.status);
    console.log(e.message);
    console.log(e instanceof NotFoundError);
    console.log(e instanceof HttpError);
    console.log(e instanceof Error);
}
"#), &["NotFoundError", "404", "User not found", "true", "true", "true"]);
}

// ===================================================================
// ERROR HANDLING PATTERNS
// ===================================================================

#[test]
fn rethrow_pattern() {
    assert_eq!(run_js(r#"
function riskyOp() {
    throw new TypeError("wrong type");
}
try {
    try {
        riskyOp();
    } catch (e) {
        if (e instanceof TypeError) {
            throw e;
        }
        console.log("handled");
    }
} catch (e) {
    console.log("rethrown: " + e.message);
}
"#), &["rethrown: wrong type"]);
}

#[test]
fn conditional_catch() {
    assert_eq!(run_js(r#"
function process(val) {
    if (val < 0) throw new RangeError("negative");
    if (val === 0) throw new TypeError("zero not allowed");
    return val * 2;
}
try {
    process(-1);
} catch (e) {
    if (e instanceof RangeError) {
        console.log("range: " + e.message);
    } else if (e instanceof TypeError) {
        console.log("type: " + e.message);
    }
}
"#), &["range: negative"]);
}

#[test]
fn try_catch_finally_order() {
    assert_eq!(run_js(r#"
let log = [];
try {
    log.push("try");
    throw new Error("oops");
} catch (e) {
    log.push("catch");
} finally {
    log.push("finally");
}
console.log(log.join(","));
"#), &["try,catch,finally"]);
}

#[test]
fn throw_string_value() {
    assert_eq!(run_js(r#"
try {
    throw "just a string";
} catch (e) {
    console.log(typeof e);
    console.log(e);
}
"#), &["string", "just a string"]);
}

#[test]
fn throw_object_value() {
    assert_eq!(run_js(r#"
try {
    throw { code: 500, msg: "internal error" };
} catch (e) {
    console.log(e.code);
    console.log(e.msg);
}
"#), &["500", "internal error"]);
}

#[test]
fn error_chaining() {
    assert_eq!(run_js(r#"
function inner() {
    throw new Error("from inner");
}
function middle() {
    try {
        inner();
    } catch (e) {
        throw new Error("from middle: " + e.message);
    }
}
try {
    middle();
} catch (e) {
    console.log(e.message);
}
"#), &["from middle: from inner"]);
}
